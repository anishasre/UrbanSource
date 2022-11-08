VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecurringBillRegisters 
   Caption         =   "frmRecurringBillRegisters"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10185
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1965
      Left            =   0
      TabIndex        =   19
      Top             =   780
      Width           =   10275
      Begin VB.CommandButton cmdSearchPaymentVoucherNo 
         Caption         =   "..."
         Height          =   270
         Left            =   8745
         TabIndex        =   13
         Top             =   150
         Width           =   285
      End
      Begin VB.TextBox txtPaymentOrderNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   2130
      End
      Begin VB.TextBox txtPaymentVoucherNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   150
         Width           =   2130
      End
      Begin VB.TextBox txtRegName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         TabIndex        =   11
         Top             =   930
         Width           =   2130
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   15
         Top             =   870
         Width           =   2130
      End
      Begin VB.TextBox txtRegID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         TabIndex        =   10
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   16
         Top             =   1230
         Width           =   2130
      End
      Begin VB.TextBox txtDemandDueDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   14
         Top             =   510
         Width           =   2130
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Voucher No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4395
         TabIndex        =   32
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Order No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Month :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5685
         TabIndex        =   25
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5715
         TabIndex        =   24
         Top             =   885
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Demand Due Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4725
         TabIndex        =   23
         Top             =   570
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "RegisterID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   22
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Register Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   21
         Top             =   960
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   10125
      TabIndex        =   0
      Top             =   0
      Width           =   10185
   End
   Begin VB.Frame Frame2 
      Height          =   2445
      Left            =   -30
      TabIndex        =   20
      Top             =   2760
      Width           =   10230
      Begin VB.TextBox txtExtraAmt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2145
         TabIndex        =   5
         Top             =   1830
         Width           =   2340
      End
      Begin VB.TextBox txtPaidAmt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Top             =   435
         Width           =   2340
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   1230
         Left            =   6210
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   825
         Width           =   3915
      End
      Begin VB.TextBox txtBilDueDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1020
         Width           =   2340
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2145
         TabIndex        =   4
         Top             =   1425
         Width           =   2340
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   255
         Width           =   2340
      End
      Begin VB.TextBox txtBillDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   615
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpBillDate 
         Height          =   315
         Left            =   4200
         TabIndex        =   18
         Top             =   615
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   40352
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Payable Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4650
         TabIndex        =   35
         Top             =   450
         Width           =   1560
      End
      Begin VB.Label Label13 
         Caption         =   "Extra Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   765
         TabIndex        =   34
         Top             =   1890
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5220
         TabIndex        =   33
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1155
         TabIndex        =   29
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Billing Due Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   28
         Top             =   1065
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Billing Date :"
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
         Left            =   540
         TabIndex        =   27
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Bill Number :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   0
         TabIndex        =   26
         Top             =   210
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10125
      TabIndex        =   30
      Top             =   5160
      Width           =   10185
      Begin VB.CommandButton cmdVerify 
         Caption         =   "test"
         Height          =   330
         Left            =   9375
         TabIndex        =   36
         Top             =   435
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmdPaymentOrder 
         Caption         =   "&Generate PaymentOrder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6120
         TabIndex        =   8
         Top             =   165
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton cmdVerifyBill 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4005
         TabIndex        =   7
         Top             =   75
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmRecurringBillRegisters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private intCheckMode As Integer      ' 1 = Normal Data Entry 2 = Past Data Entry
    Dim mVoucherID As Variant
    Dim mInstrumentTypeID As Variant
    Dim mInstrumentNo As Variant
    Dim mInstrumentDate As Variant
    Dim mintVoucherNo As Variant
    Dim mArrIn As Variant
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mPaymentOrderID As Variant
    Dim mPaymentOrderNo As Variant
    Dim mID As Integer
    Dim mPeriodID As Variant
    Dim mMonthID As Variant
    Dim objBk As New clsBank
  


    Private Sub cmdSearchPaymentVoucherNo_Click()
        Dim mcnn  As New ADODB.Connection
        Dim Rec   As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL  As String
        
        frmSearchVouchers.CheckMode = 20
        frmSearchVouchers.chkContra.Visible = False
        frmSearchVouchers.chkReceipt.Visible = False
        frmSearchVouchers.chkJournal.Visible = False
        frmSearchVouchers.chkPayment.value = 1
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            txtPaymentVoucherNo.Text = gbSearchCode
            txtPaymentVoucherNo.Tag = gbSearchID
'            gbSearchCode = ""
'            gbSearchID = -1
        End If
        If gbSearchCode <> "" Then
            If objDB.SetConnection(mcnn) Then
                mSQL = " SELECT  fltAmount,intKeyID2, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, intVoucherNo"
                mSQL = mSQL + " FROM faVouchers Where    (NOT (intKeyID2 IS NULL)) and intVoucherNo = " & gbSearchCode & " "
                Rec.Open mSQL, mcnn
                If Not (Rec.EOF And Rec.BOF) Then
                    txtPaymentOrderNo.Text = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                    mInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    mInstrumentNo = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    mInstrumentDate = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    txtPaidAmt.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                End If
                Rec.Close
              Else
                MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
            End If
            gbSearchCode = ""
            gbSearchID = -1
       End If
        
    End Sub

    Private Sub cmdPaymentOrder_Click()
'        If SaveValidation = False Then Exit Sub
'        If intCheckMode = 1 Then
'            Call SavePaymentOrder
'        ElseIf intCheckMode = 2 Then
'            Call SavePastData
'        End If
    End Sub

    Private Sub cmdVerifyBill_Click()
        Dim objDB As New clsDB
        Dim mcnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mintID As Variant
        Dim mSQL As String
        Dim mAccountHeadCode As Variant
        Dim mAccountHeadID As Variant
        Dim mFunctionID As Integer
        Dim mFunctionaryID As Integer
        Dim mPaidAmt As Variant
        
        If SaveValidation = False Then Exit Sub
        'mPaidAmt = val(txtAmount.Text) + val(txtExtraAmt.Text)
        
        objDB.SetConnection mcnn
        mSQL = "SELECT faRegisterOfBills.intFunctionaryID, faRegisterOfBills.intFunctionID, faAccountHeads.vchAccountHeadCode,faAccountHeads.intAccountHeadID,faBillRegisters.tnyPreriodID"
        mSQL = mSQL + " FROM faBillRegisters INNER JOIN "
        mSQL = mSQL + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID INNER JOIN"
        mSQL = mSQL + " faAccountHeads ON faRegisterOfBills.intExpenditureHeadID = faAccountHeads.intAccountHeadID"
        mSQL = mSQL + " Where faBillRegisters.intID = " & txtRegID.Tag & " "
        
        Rec.Open mSQL, mcnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
       
        If Not (Rec.EOF And Rec.BOF) Then
            mAccountHeadCode = Rec!vchAccountHeadCode
            mAccountHeadID = Rec!intAccountHeadID
            mFunctionID = Rec!intFunctionID
            mFunctionaryID = Rec!intFunctionaryID
            mPeriodID = IIf(IsNull(Rec!tnyPreriodID), Null, Rec!tnyPreriodID)
        End If
        
        If Trim(txtMonth.Text) = "January" Then
            mMonthID = 1
        ElseIf Trim(txtMonth.Text) = "February" Then
            mMonthID = 2
        ElseIf Trim(txtMonth.Text) = "March" Then
            mMonthID = 3
        ElseIf Trim(txtMonth.Text) = "April" Then
            mMonthID = 4
        ElseIf Trim(txtMonth.Text) = "May" Then
            mMonthID = 5
        ElseIf Trim(txtMonth.Text) = "June" Then
            mMonthID = 6
        ElseIf Trim(txtMonth.Text) = "July" Then
            mMonthID = 7
        ElseIf Trim(txtMonth.Text) = "August" Then
            mMonthID = 8
        ElseIf Trim(txtMonth.Text) = "September" Then
            mMonthID = 9
        ElseIf Trim(txtMonth.Text) = "October" Then
            mMonthID = 10
        ElseIf Trim(txtMonth.Text) = "November" Then
            mMonthID = 11
        ElseIf Trim(txtMonth.Text) = "December" Then
            mMonthID = 12
        End If
   
         mID = IIf(txtRegID.Tag = "", -1, val(txtRegID.Tag))
         mArrIn = Array(mID, txtDemandDueDate.Text, _
                            txtRegID.Text, _
                            txtYear.Text, _
                            mMonthID, _
                            mPeriodID, _
                            Trim(txtBillNo.Text), _
                            Trim(txtBillDate.Text), _
                            Trim(txtBilDueDate.Text), _
                            Trim(txtAmount.Text), _
                            Trim(txtPaidAmt.Text), _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Trim(txtRemarks.Text), _
                            2, _
                            Null _
                            )
         objDB.ExecuteSP "spSaveBillRegisters", mArrIn, , , mcnn, adCmdStoredProc
         MsgBox "Updated Successfully", vbInformation, "Saankhya"
         cmdVerifyBill.Enabled = False
         'frmListofBillRegister.CheckDemandID = 1
         Unload Me
         'frmListOfRegisterOfBills.Show
    End Sub

    Private Sub dtpBillDate_CloseUp()
        txtBillDate.Text = DdMmmYy(dtpBillDate.value)
    End Sub

    Private Sub Form_Load()
        Call FormInitialize
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        
        If intCheckMode = 1 Then
             Frame1.Enabled = False
             txtPaymentOrderNo.Locked = True
             txtPaymentVoucherNo.Locked = True
        ElseIf intCheckMode = 2 Then
             Frame1.Enabled = True
             cmdPaymentOrder.Visible = True
             txtPaymentOrderNo.Locked = True
             txtPaymentVoucherNo.Locked = True
             txtRegID.Enabled = False
             txtRegName.Enabled = False
             txtMonth.Enabled = False
             txtYear.Enabled = False
             txtDemandDueDate.Enabled = False
        End If
        txtExtraAmt.Enabled = False
        txtAmount.Text = Format(Abs(objBk.Opening), "0.00")
        txtExtraAmt.Text = Format(Abs(objBk.Opening), "0.00")
        txtPaidAmt.Text = Format(Abs(objBk.Opening), "0.00")
        'cmdSearchPaymentVoucherNo.Enabled = False
        cmdSearchPaymentVoucherNo.Enabled = True
        dtpBillDate.value = gbTransactionDate
    End Sub
    Private Function SaveValidation() As Boolean
        If Trim(txtBillNo.Text) = "" Then
             MsgBox "Enter the Bill Number", vbInformation
             txtBillNo.SetFocus
             SaveValidation = False
             Exit Function
        End If
        If Trim(txtBillDate.Text) = "" Then
            MsgBox "Enter the Bill Date", vbInformation
            txtBillDate.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtBilDueDate.Text) = "" Then
            MsgBox "Enter the Bill Due Date", vbInformation
            txtBilDueDate.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtAmount.Text) = "" Or Trim(txtAmount.Text) <= 0 Then
            MsgBox "Enter the Bill Amount", vbInformation
            txtAmount.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtPaidAmt.Text) = "" Then
            MsgBox "Enter the Paid Amount", vbInformation
            txtPaidAmt.SetFocus
            SaveValidation = False
            Exit Function
        End If
        
        If Trim(txtRemarks.Text) = "" Then
            MsgBox "Enter Remarks", vbInformation
            txtRemarks.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If txtBillDate.Text > txtBilDueDate.Text Then
            If Trim(txtExtraAmt.Text) = "0.00" Then
                MsgBox "Enter Extra Amount", vbInformation
                txtExtraAmt.SetFocus
                SaveValidation = False
                Exit Function
            End If
        End If
        SaveValidation = True
    End Function

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    Private Sub txtBilDate_LostFocus()
        txtBillDate.Text = CheckDateInMMM(txtBillDate.Text)
        
    End Sub
    Private Sub txtAmount_LostFocus()
    txtPaidAmt.Text = txtAmount.Text
        If txtBillDate.Text > txtBilDueDate.Text Then
            MsgBox "Enter extra amount please", vbInformation
            txtExtraAmt.Enabled = True
            txtExtraAmt.SetFocus
        End If
        
    End Sub

    Private Sub txtBilDueDate_LostFocus()
        txtBilDueDate = CheckDateInMMM(txtBilDueDate.Text)
        
    End Sub

    Private Sub txtBillDate_LostFocus()
        txtBillDate.Text = CheckDateInMMM(txtBillDate.Text)
    End Sub

    Private Sub txtDemandDueDate_LostFocus()
        txtDemandDueDate.Text = CheckDateInMMM(txtDemandDueDate.Text)
    End Sub



    Private Sub txtExtraAmt_LostFocus()
           txtPaidAmt.Text = val(txtExtraAmt.Text) + val(txtAmount.Text)
    End Sub

    Private Sub txtPaidAmt_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
   
'''    Private Sub SavePaymentOrder()
'''        Dim PO As uPaymentOrder
'''        Dim POC As uPaymentOrderChild
'''        Dim POAdd As uPaymentOrderAddress
'''        Dim objDb As New clsDb
'''        Dim mCnn As New ADODB.Connection
'''        Dim Rec As New ADODB.Recordset
'''        Dim Rec1 As New ADODB.Recordset
'''        Dim mSLNo As Integer
'''        Dim mLoop As Integer
'''        Dim vchPayOrderNo As String
'''        Dim mintID As Variant
'''        Dim mSql As String
'''        Dim mAccountHeadCode As Variant
'''        Dim mAccountHeadID As Variant
'''        Dim mFunctionID As Integer
'''        Dim mFunctionaryID As Integer
'''
'''
'''        cmdVerify.Enabled = False
'''
''''-------------------------to generate Voucher Number---------------------------------------'
'''
'''        objDb.SetConnection mCnn
'''        mSql = "SELECT faRegisterOfBills.intFunctionaryID, faRegisterOfBills.intFunctionID, faAccountHeads.vchAccountHeadCode,faAccountHeads.intAccountHeadID,faBillRegisters.tnyPreriodID"
'''        mSql = mSql + " FROM faBillRegisters INNER JOIN "
'''        mSql = mSql + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID INNER JOIN"
'''        mSql = mSql + " faAccountHeads ON faRegisterOfBills.intExpenditureHeadID = faAccountHeads.intAccountHeadID"
'''        mSql = mSql + " Where faBillRegisters.intID = " & txtRegID.Tag & " "
'''
'''        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
'''
'''        If Not (Rec.EOF And Rec.BOF) Then
'''            mAccountHeadCode = Rec!vchAccountHeadCode
'''            mAccountHeadID = Rec!intAccountHeadID
'''            mFunctionID = Rec!intFunctionID
'''            mFunctionaryID = Rec!intFunctionaryID
'''            mPeriodID = IIf(IsNull(Rec!tnyPreriodID), Null, Rec!tnyPreriodID)
'''        End If
'''        arrInput = Array(mAccountHeadCode, _
'''                                    20, _
'''                                    gbFinancialYearID, _
'''                                    mintVoucherNo _
'''                         )
'''        Rec.Close
'''
'''        mSql = "Declare @intVoucherNo Numeric " & vbNewLine
'''        mSql = mSql + " Exec spGetVoucherNo Null,20," & gbFinancialYearID & ",@intVoucherNo output" & vbNewLine
'''        mSql = mSql + " Select @intVoucherNo [numVoucherNo]"
'''        Rec.Open mSql, mCnn
'''        mintVoucherNo = Rec!numVoucherNo
'''        'Rec.Close
'''
''''------------------------------------------------------------------------------------------------'
''''        If val(txtPaymentOrderNo.Tag) <> 0 Then
''''            mSql = "Select tnyStatus From faPayOrder Where intPayOrderID = " & val(txtPaymentOrderNo.Tag)
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF Or Rec.BOF) Then
''''                If Rec!tnyStatus <> 0 Then
''''                    MsgBox "Sorry! This Payment Order is already approved. Editing is not Permitted", vbInformation
''''                    Exit Sub
''''                End If
''''            End If
''''            mSql = ""
''''            If Rec.State = 1 Then Rec.Close
''''        End If
''''----------------------------------to generate Payment Order Number-------------------------------------'
'''
'''        With PO
'''            .intPayOrderID = IIf(txtPaymentOrderNo.Tag = "", Null, txtPaymentOrderNo.Tag)
'''            .vchPayOrderNo = IIf(txtPaymentVoucherNo.Text = "", Null, txtPaymentVoucherNo.Text)
'''            .dtPayOrderDate = Trim(txtBillDate.Text)
'''            .dtDueDate = Trim(txtBilDueDate.Text)
'''            .intFunctionaryID = mFunctionaryID
'''            .intFunctionID = mFunctionID
'''            .intTransactionTypeID = Null
'''            .vchBillNo = Trim(txtBillNo.Text)
'''            .numBillAmount = Trim(txtAmount.Text)
'''            .dtBillDate = Trim(txtBillDate.Text)
'''            .intInstrumentTypeID = Null
'''            .intCashOrBankHeadID = Null
'''            .vchDescription = Trim(txtRemarks.Text)
'''            .vchTitle = Null
'''            .intSubLedgerTypeID = Null
'''            .intPayToSubLedgerID = Null
'''            .intSubsidiaryCashBookID = Null
'''            .intImplementingOfficerID = Null
'''            .numProjectNo = Null
'''            .intStockRegisterID = Null
'''            .vchStockRefNo = Null
'''            .intAssetTypeID = Null
'''            .intAssetID = Null
'''            .numFwdSeatID = Null
'''            .intLocalBodyID = gbLocalBodyID
'''            .intZonalID = gbLocationID
'''            .intFinancialYearID = gbFinancialYearID
'''            .numUserID = gbUserID
'''            .numSeatID = gbSeatID
'''            .numApprovingOfficerID = Null
'''            .numApprovingSeatID = Null
'''            .dtApprovingDate = Null
'''            .intSourceOfFundID = Null
'''            .intAllotmentID = Null
'''            .intAgreementID = Null
'''            .tnyCategoryID = Null
'''            .tnySectorID = Null
'''            .tnyIsFinalBill = Null
'''            .intVoucherID = Null
'''            .intVoucherNo = mintVoucherNo
'''            .dtVoucherDate = Null
'''            .tnyStatus = 0
'''            .intKeyID = Null 'Section ID stores from Pay Bill -Sthapana for Pay&Allowance
'''            .numKeyID = Null
'''            .dtKeyDate = Null
'''            .tnyCancelled = 0
'''            .intAppID = 115
'''            .intModuleID = Null
'''
'''            arrInput = Array(.intPayOrderID, .vchPayOrderNo, .dtPayOrderDate, .dtDueDate, .intFunctionaryID, _
'''            .intFunctionID, .intTransactionTypeID, .vchBillNo, .numBillAmount, _
'''            .dtBillDate, .intInstrumentTypeID, .intCashOrBankHeadID, .vchDescription, _
'''            .vchTitle, .intSubLedgerTypeID, .intPayToSubLedgerID, .intSubsidiaryCashBookID, _
'''            .intImplementingOfficerID, .numProjectNo, .intStockRegisterID, .vchStockRefNo, _
'''            .intAssetTypeID, .intAssetID, .numFwdSeatID, .intLocalBodyID, _
'''            .intZonalID, .intFinancialYearID, .numUserID, .numSeatID, _
'''            .numApprovingOfficerID, .numApprovingSeatID, .dtApprovingDate, .intVoucherID, .intVoucherNo, .dtVoucherDate, _
'''            .tnyStatus, .intKeyID, .numKeyID, .dtKeyDate, .tnyCancelled, .intAppID, .intModuleID, .intSourceOfFundID, _
'''            .intAllotmentID, .intAgreementID, .tnyCategoryID, .tnySectorID, _
'''            .tnyIsFinalBill)
'''
'''            objDb.ExecuteSP "spSavePayOrder", arrInput, arrOutPut, , mCnn, adCmdStoredProc
'''        End With
'''        If IsNumeric(arrOutPut(0, 0)) Then
'''            mPaymentOrderID = arrOutPut(0, 0)
'''            vchPayOrderNo = arrOutPut(1, 0)
'''        End If
'''        Rec.Close
'''        '---------------------------------------------------------------------------------------------------------
'''
'''        mSql = "Delete From faPayOrderChild Where intPayOrderID = " & mPaymentOrderID
'''        mCnn.Execute mSql
'''With frmListofBillRegister.vsGrid
''' For mLoop = 1 To .Rows - 1
'''   If .Cell(flexcpChecked, mLoop, 6) = 1 Then
'''     mSql = " SELECT * From faBillRegisters Where (tnyStatus = 2)AND intID =" & .TextMatrix(mLoop, 7) & " "
'''     Rec1.Open mSql, mCnn
'''
'''     If (Rec1.EOF And Rec1.BOF) Then
'''       mSql = "SELECT faRegisterOfBills.intExpenditureHeadID, faBillRegisters.tnyStatus, faBillRegisters.intID"
'''       mSql = mSql + " FROM faRegisterOfBills INNER JOIN faBillRegisters ON faRegisterOfBills.intRegID = faBillRegisters.intRegID "
'''       mSql = mSql + " WHERE  faBillRegisters.intID = " & .TextMatrix(mLoop, 7) & " "
'''       Rec.Open mSql, mCnn
'''       txtRegID.Tag = Rec!intID
'''
'''        mSLNo = mSLNo + 1
'''        With POC
'''            .intPayOrderID = mPaymentOrderID
'''            .intSlNo = mSLNo
'''            .intAccountHeadID = mAccountHeadID
'''            .vchAccountHeadCode = mAccountHeadCode
'''            .numAmount = val(txtAmount.Text)
'''            .tnyCategoryFlag = 1
'''            .tnyDebitOrCreditFlag = 1
'''            .vchDescription = Null
'''
'''            arrInput = Array(.intPayOrderID, _
'''            .intSlNo, _
'''            .intAccountHeadID, _
'''            .vchAccountHeadCode, _
'''            .numAmount, _
'''            .tnyCategoryFlag, _
'''            .tnyDebitOrCreditFlag, _
'''            .vchDescription)
'''
'''            objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
'''        End With
''''---------------------------------------------------------------------------------------------------------------'
''''                                        PAYORDER CHILD                                                                       '
''''---------------------------------------------------------------------------------------------------------------'
''''        mSLNo = mSLNo + 1
''''        'For mLoop = 1 To vsGrid.Rows - 1
''''            'If val(vsGrid.TextMatrix(mLoop, 1)) > 0 And val(vsGrid.TextMatrix(mLoop, 3)) > 0 Then
''''            With POC
''''                .intPayOrderID = mPaymentOrderID
''''                .intSlNo = mSLNo
''''                .intAccountHeadID = mAccountHeadID
''''                .vchAccountHeadCode = mAccountHeadCode
''''                .numAmount = val(txtAmount.Text)
''''                .tnyCategoryFlag = 2
''''                .tnyDebitOrCreditFlag = 0
''''                .vchDescription = Null
''''
''''                arrInput = Array(.intPayOrderID, _
''''                .intSlNo, _
''''                .intAccountHeadID, _
''''                .vchAccountHeadCode, _
''''                .numAmount, _
''''                .tnyCategoryFlag, _
''''                .tnyDebitOrCreditFlag, _
''''                .vchDescription)
''''                objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
''''            End With
''''            'End If
''''        'Next
''''
''''        mSLNo = mSLNo + 1
''''        With POC
''''            .intPayOrderID = mPaymentOrderID
''''            .intSlNo = mSLNo
''''            .intAccountHeadID = mAccountHeadID
''''            .vchAccountHeadCode = mAccountHeadCode
''''            .numAmount = val(txtAmount.Text)
''''            .tnyCategoryFlag = 3
''''            .tnyDebitOrCreditFlag = 0
''''            .vchDescription = Null
''''
''''            arrInput = Array(.intPayOrderID, _
''''            .intSlNo, _
''''            .intAccountHeadID, _
''''            .vchAccountHeadCode, _
''''            .numAmount, _
''''            .tnyCategoryFlag, _
''''            .tnyDebitOrCreditFlag, _
''''            .vchDescription)
''''            objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
''''        End With
''''
''''        If val(txtBillNo.Text) <> 0 Then
''''            mSLNo = mSLNo + 1
''''            With POC
''''                .intPayOrderID = mPaymentOrderID
''''                .intSlNo = mSLNo
''''                .intAccountHeadID = mAccountHeadID
''''                .vchAccountHeadCode = mAccountHeadCode
''''                .numAmount = val(txtAmount.Text)
''''                .tnyCategoryFlag = 4
''''                .tnyDebitOrCreditFlag = 0
''''                .vchDescription = Null
''''
''''                arrInput = Array(.intPayOrderID, _
''''                .intSlNo, _
''''                .intAccountHeadID, _
''''                .vchAccountHeadCode, _
''''                .numAmount, _
''''                .tnyCategoryFlag, _
''''                .tnyDebitOrCreditFlag, _
''''                .vchDescription)
''''                objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
''''            End With
''''        End If
'''
'''
'''
''''---------------------------------------------------------------------------------------------------'
''''                                    TO SAVE IN faBillRegister                                      '                         '
''''---------------------------------------------------------------------------------------------------'
'''
'''        If Trim(txtMonth.Text) = "January" Then
'''            mMonthID = 1
'''        ElseIf Trim(txtMonth.Text) = "February" Then
'''            mMonthID = 2
'''        ElseIf Trim(txtMonth.Text) = "March" Then
'''            mMonthID = 3
'''        ElseIf Trim(txtMonth.Text) = "April" Then
'''            mMonthID = 4
'''        ElseIf Trim(txtMonth.Text) = "May" Then
'''            mMonthID = 5
'''        ElseIf Trim(txtMonth.Text) = "June" Then
'''            mMonthID = 6
'''        ElseIf Trim(txtMonth.Text) = "July" Then
'''            mMonthID = 7
'''        ElseIf Trim(txtMonth.Text) = "August" Then
'''            mMonthID = 8
'''        ElseIf Trim(txtMonth.Text) = "September" Then
'''            mMonthID = 9
'''        ElseIf Trim(txtMonth.Text) = "October" Then
'''            mMonthID = 10
'''        ElseIf Trim(txtMonth.Text) = "November" Then
'''            mMonthID = 11
'''        ElseIf Trim(txtMonth.Text) = "December" Then
'''            mMonthID = 12
'''        End If
'''
'''         mID = IIf(txtRegID.Tag = "", -1, val(txtRegID.Tag))
'''         mArrIn = Array(mID, txtDemandDueDate.Text, _
'''                            txtRegID.Text, _
'''                            txtYear.Text, _
'''                            mMonthID, _
'''                            mPeriodID, _
'''                            Trim(txtBillNo.Text), _
'''                            Trim(txtBillDate.Text), _
'''                            Trim(txtBilDueDate.Text), _
'''                            Trim(txtAmount.Text), _
'''                            Trim(txtAmount.Text), _
'''                            vchPayOrderNo, _
'''                            mintVoucherNo, _
'''                            Null, _
'''                            Null, _
'''                            Null, _
'''                            Trim(txtRemarks.Text), _
'''                            2 _
'''                            )
'''         objDb.ExecuteSP "spSaveBillRegisters", mArrIn, , , mCnn, adCmdStoredProc
'''
'''            End If
'''           ' Rec.Close
'''        End If
''''        Rec1.Close
'''    Next mLoop
'''End With
'''
''''-------------------------------------------------------------------------------------------------------------------'
''''                                           PAYORDER ADDRESS                                                        '
''''-------------------------------------------------------------------------------------------------------------------'
'''
'''    With POAdd
'''
'''            'ObjSubLed.SetSubLedgerDetails (val(txtName.Tag))
'''            .intPayOrderID = mPaymentOrderID
'''            .intSubsidiaryAccountHeadID = Null
'''            .intSubLegerTypeID = Null
'''            .vchSubLedgerCode = Null
'''            .vchName = Null
'''            .vchHouseName = Null
'''            .vchStreet = Null
'''            .vchLocalPlace = Null
'''            .vchMainPlace = Null
'''            .vchPost = Null
'''            .vchPinCode = Null
'''            .vchPhone = Null
'''
'''            arrInput = Array(.intPayOrderID, _
'''            .intSubsidiaryAccountHeadID, _
'''            .intSubLegerTypeID, _
'''            .vchSubLedgerCode, _
'''            .vchName, _
'''            .vchHouseName, _
'''            .vchStreet, _
'''            .vchLocalPlace, _
'''            .vchMainPlace, _
'''            .vchPost, _
'''            .vchPinCode, _
'''            .vchPhone)
'''            objDb.ExecuteSP "spSavePayOrderAddress", arrInput, , , mCnn, adCmdStoredProc
'''
'''        End With
''''----------------------------------------------------------------------------------------------------------------
'''
'''
'''
'''         MsgBox "Saved Payment!", vbInformation, "Saankhya"
'''         txtPaymentOrderNo.Text = vchPayOrderNo
'''         txtPaymentVoucherNo.Text = mintVoucherNo
'''         txtPaidAmt.Text = txtAmount
'''         cmdVerify.Enabled = False
'''  End Sub
    
''    Public Property Let CheckMode(mData As Integer)
''        intCheckMode = mData
''    End Property
''    Public Property Get CheckMode() As Integer
''        CheckMode = intCheckMode
''    End Property
''    Private Sub SavePastData()
''    Dim ObjDb As New clsDb
''    Dim mCnn As New ADODB.Connection
''    Dim Rec As New ADODB.Recordset
''    Dim msQl As String
''    Dim mID As Integer
''    Dim mMonthID As Variant
''
''        ObjDb.SetConnection mCnn
''        msQl = "SELECT tnyPreriodID From faBillRegisters Where intID = " & txtRegID.Tag & " "
''        Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''
''        If Not (Rec.EOF And Rec.BOF) Then
''             mPeriodID = IIf(IsNull(Rec!tnyPreriodID), Null, Rec!tnyPreriodID)
''        Else
''             mPeriodID = Null
''        End If
''        Rec.Close
''        If Trim(txtMonth.Text) = "January" Then
''            mMonthID = 1
''        ElseIf Trim(txtMonth.Text) = "February" Then
''            mMonthID = 2
''        ElseIf Trim(txtMonth.Text) = "March" Then
''            mMonthID = 3
''        ElseIf Trim(txtMonth.Text) = "April" Then
''            mMonthID = 4
''        ElseIf Trim(txtMonth.Text) = "May" Then
''            mMonthID = 5
''        ElseIf Trim(txtMonth.Text) = "June" Then
''            mMonthID = 6
''        ElseIf Trim(txtMonth.Text) = "July" Then
''            mMonthID = 7
''        ElseIf Trim(txtMonth.Text) = "August" Then
''            mMonthID = 8
''        ElseIf Trim(txtMonth.Text) = "September" Then
''            mMonthID = 9
''        ElseIf Trim(txtMonth.Text) = "October" Then
''            mMonthID = 10
''        ElseIf Trim(txtMonth.Text) = "November" Then
''            mMonthID = 11
''        ElseIf Trim(txtMonth.Text) = "December" Then
''            mMonthID = 12
''        End If
''
''         mID = IIf(txtRegID.Tag = "", -1, val(txtRegID.Tag))
''         mArrIn = Array(mID, txtDemandDueDate.Text, _
''                            txtRegID.Text, _
''                            txtYear.Text, _
''                            mMonthID, _
''                            mPeriodID, _
''                            Trim(txtBillNo.Text), _
''                            Trim(txtBillDate.Text), _
''                            Trim(txtBilDueDate.Text), _
''                            Trim(txtAmount.Text), _
''                            Trim(txtPaidAmt.Text), _
''                            Trim(txtPaymentOrderNo.Text), _
''                            Trim(txtPaymentVoucherNo.Text), _
''                            mInstrumentTypeID, _
''                            mInstrumentNo, _
''                            mInstrumentDate, _
''                            Trim(txtRemarks.Text), _
''                            2 _
''                            )
''         ObjDb.ExecuteSP "spSaveBillRegisters", mArrIn, , , mCnn, adCmdStoredProc
''         MsgBox "Saved Payment!", vbInformation, "Saankhya"
''         cmdPaymentOrder.Enabled = False
''    End Sub



