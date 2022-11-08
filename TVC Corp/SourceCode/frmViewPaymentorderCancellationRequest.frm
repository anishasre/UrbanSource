VERSION 5.00
Begin VB.Form frmViewPaymentorderCancellationRequest 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paymentorder Cancellation Request View"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
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
   ScaleHeight     =   7095
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewPO 
      Caption         =   "..."
      Height          =   330
      Left            =   5355
      TabIndex        =   23
      Top             =   585
      Width           =   420
   End
   Begin VB.TextBox txtApprover 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3300
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   4035
      TabIndex        =   1
      Top             =   6210
      Width           =   1320
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   420
      Left            =   2640
      TabIndex        =   0
      Top             =   6210
      Width           =   1320
   End
   Begin VB.TextBox txtCancelledDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5535
      Width           =   3300
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4590
      Width           =   3300
   End
   Begin VB.TextBox txtForwardedSeat 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3825
      Width           =   3300
   End
   Begin VB.TextBox txtSeat 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3375
      Width           =   3300
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3300
   End
   Begin VB.TextBox txtRequestDate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2385
      Width           =   3300
   End
   Begin VB.TextBox txtReson 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1035
      Width           =   3300
   End
   Begin VB.TextBox txtPayOrderNo 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   540
      Width           =   3300
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   2010
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1485
      Width           =   3300
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   795
      TabIndex        =   22
      Top             =   4320
      Width           =   1290
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approver"
      Height          =   270
      Left            =   1200
      TabIndex        =   20
      Top             =   5130
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelled Date"
      Height          =   270
      Left            =   660
      TabIndex        =   10
      Top             =   5580
      Width           =   1245
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   270
      Left            =   1380
      TabIndex        =   9
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarded Seat"
      Height          =   270
      Left            =   615
      TabIndex        =   8
      Top             =   3915
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
      Height          =   270
      Left            =   795
      TabIndex        =   7
      Top             =   2430
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seat"
      Height          =   270
      Left            =   1515
      TabIndex        =   6
      Top             =   3465
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   270
      Left            =   1020
      TabIndex        =   5
      Top             =   2925
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   270
      Left            =   1290
      TabIndex        =   4
      Top             =   1530
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      Height          =   270
      Left            =   1335
      TabIndex        =   3
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Order No."
      Height          =   270
      Left            =   345
      TabIndex        =   2
      Top             =   585
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      Height          =   6675
      Left            =   180
      Top             =   180
      Width           =   6315
   End
End
Attribute VB_Name = "frmViewPaymentorderCancellationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPaymentOrderNo As Variant
Private mUserTypeID As Integer
Private mArrayIn As Variant
Private mVerified As Boolean

    Public Property Get PaymentOrderNo() As Variant
        PaymentOrderNo = mPaymentOrderNo
    End Property
    
    Public Property Let PaymentOrderNo(ByVal argc As Variant)
        mPaymentOrderNo = argc
    End Property
    
    Public Property Let UserType(argc As Integer)
        mUserTypeID = argc
    End Property
    
    Public Property Let Verified(Data As Boolean)
        mVerified = Data
    End Property
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdView_Click()
'        If mVerified Then
'            Dim mCnn As New ADODB.Connection
'            Dim Rec As New ADODB.Recordset
'            Dim objDB As New clsDB
'            Dim mSQL As String
'            Dim mArrayOut As Variant
'
'            If gbSectionID <> 4 Then
'                MsgBox "Accounts Section can only Apply for cancellation", vbInformation
'                Exit Sub
'            End If
'            If txtStatus.Tag <> 0 Then
'                MsgBox "Already Cancelled", vbInformation
'                Exit Sub
'            End If
'
'            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
'            Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where tnyVoucherTypeID = 20 And intKeyID2 = " & mPaymentOrderNo, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'
'            End If
'
'            If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then 'For an Appoving Officer
'                '-------------------------------------------------------------------------------------'
'                '                                       Reversing                                     '
'                '-------------------------------------------------------------------------------------'
'                Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPaymentOrderNo, mCnn
'                If Not (Rec.EOF And Rec.BOF) Then
'                    While Not (Rec.EOF)
'                        mArrayIn = Array(Rec!intVoucherID, gbTransactionDate)
'                        objDB.ExecuteSP "spSaveReverseVouchers", mArrayIn, mArrayOut, , mCnn
'                        If Rec!tnyVoucherTypeID = 20 Then
'                            mCnn.Execute "Update faVouchers Set tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
'                            mCnn.Execute "Update faVouchers Set tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
'                        End If
'                        Rec.MoveNext
'                    Wend
'                Else
'                    MsgBox "No Journal to Reverse", vbInformation
'                End If
'                Rec.Close
'                '-------------------------------------------------------------------------------------'
'
'                mSQL = "Update faCancelledVouchers Set tnyApproveStatus = 1,dtCancellationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',intApproverID = " & gbUserID & " Where numReceiptNo = " & mPaymentOrderNo
'                mCnn.Execute mSQL
'                mCnn.Execute "Update faPayOrder Set tnyCancelled = 1 Where vchPayorderNo = " & mPaymentOrderNo
'            End If
'            Call ViewCancelDetails
'        Else
'            frmViewPayorderCancelRequest.ArrayIn = mArrayIn
'            Unload Me
'            frmViewPayorderCancelRequest.Visible = True
'            frmViewPayorderCancelRequest.ZOrder (0)
'        End If
        
        
        '' New Corrections ''
        If mVerified Then
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            Dim mSQL As String
            Dim mArrayOut As Variant
            Dim mPaymentFlag As Boolean
            Dim mApproveStatus As Integer
            
            mPaymentFlag = False
            mApproveStatus = 0

            If gbSectionID <> 4 And gbSeatGroupID <> gbSeatGroupSecretary Then
                MsgBox "Accounts Section can only Apply for cancellation", vbInformation
                Exit Sub
            End If
            If txtStatus.Tag = 2 Then
                MsgBox "Already Cancelled", vbInformation
                Exit Sub
            End If

            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            '-------------------------------------------------'
            '''         Checking Approve Status             '''
            '-------------------------------------------------'
            Rec.Open "SELECT tnyApproveStatus FROM faCancelledVouchers WHERE numReceiptNo = " & mPaymentOrderNo, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If IsNull(Rec!tnyApproveStatus) = False Then
                    mApproveStatus = Rec!tnyApproveStatus
                    If mApproveStatus = 2 Then
                        MsgBox "Already Cancelled", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            Rec.Close
            '---------------------------------------------------'
            
            '''         Checking if Payment already Made        '''
            '---------------------------------------------------'
            Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where tnyVoucherTypeID = 20 And intKeyID2 = " & mPaymentOrderNo, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mPaymentFlag = True
            End If
            Rec.Close
            '---------------------------------------------------'
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then 'For an Appoving Officer
            '''     If Payment not Done
                If mPaymentFlag = False Then
                '''     Reverse All,Clear Reconcile Flag, faCancelledVouchers.ApproveStatus = 2, faPayOrder.tnyCancelled = 1
                    Call CancelPayOrder(mCnn)
                Else
                    mCnn.Execute "Update faCancelledVouchers Set tnyApproveStatus = 1 Where numReceiptNo = " & mPaymentOrderNo
                '''     faCancelled Vouchers ApproveStatus = 1
                End If
            ElseIf gbSeatGroupID = gbSeatGroupSecretary Then        '   Secretary
                '''     Check ApproveStatus = 1 Otherwise Message
                If mApproveStatus = 1 Then
                '''     Reverse All,Clear Reconcile Flag, faCancelledVouchers.ApproveStatus = 2, faPayOrder.tnyCancelled = 1
                    Call CancelPayOrder(mCnn)
                Else
                    MsgBox "Approval of payment order from Accounts Supdt. not Completed", vbInformation
                End If
            Else
                MsgBox "Invalid Seat Group"
            End If
            Call ViewCancelDetails
        Else
            frmViewPayorderCancelRequest.ArrayIn = mArrayIn
            Unload Me
            frmViewPayorderCancelRequest.Visible = True
            frmViewPayorderCancelRequest.ZOrder (0)
        End If
    End Sub

    Private Sub cmdViewPO_Click()
        Dim aryIn As Variant
        aryIn = Array(mPaymentOrderNo)
        frmViewVoucher.ArrayIn = aryIn
        frmViewVoucher.FormName = "frmViewPaymentOrder"
        frmViewVoucher.Show vbModal
    End Sub

    Private Sub Form_Activate()
        Me.Top = (frmMenu.Height - Me.Height) / 2
        Me.Left = (frmMenu.Width - Me.Width) / 2
        
        Call ViewCancelDetails
    End Sub
    Private Sub ViewCancelDetails()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSQL As String
        Dim mCnt As Double
        mCnt = 1
        txtPayOrderNo.Tag = -1
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mArrayIn = Array(mUserTypeID, gbUserID, PaymentOrderNo)
        Set Rec = objdb.ExecuteSP("spGetListofCancelledPayments", mArrayIn, , , mCnn)
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                txtPayOrderNo.Tag = Rec!intVoucherID
                txtPayOrderNo.Text = Rec!numReceiptNo
                txtRequestDate.Text = Format(Rec!dtRequestDate, "dd/MMM/YYYY")
                txtUserName.Tag = Rec!intUserID
                txtUserName.Text = Rec!vchUserName
                txtReson.Tag = Rec!intReasonID
                txtReson.Text = Rec!vchCancelReason
                txtApprover.Tag = Rec!numApproverUserID
                txtApprover.Text = Rec!ApproverName
                If IsNull(Rec!dtCancellationDate) Then
                    txtCancelledDate.Text = ""
                Else
                    txtCancelledDate.Text = Format(Rec!dtCancellationDate, "dd/MMM/YYYY")
                End If
                txtStatus.Tag = Rec!Status
                If Rec!Status = 0 Then
                    txtStatus.Text = "Not Cancelled"
                ElseIf Rec!Status = 1 Then
                    txtStatus.Text = "First Level Approval Completed"
                Else
                    txtStatus.Text = "Cancelled"
                End If
                txtForwardedSeat.Text = Rec!chvSeatTitle
                txtSeat.Text = gbSeatName
                txtRemarks.Text = Rec!vchRemarks
                Rec.MoveNext
            Wend
        End If
    End Sub
    Private Sub CancelPayOrder(mCnn As ADODB.Connection)
        Dim mArrayOut As Variant
        Dim mSQL As String
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        '-------------------------------------------------------------------------------------'
        '                                       Reversing                                     '
        '-------------------------------------------------------------------------------------'
        Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPaymentOrderNo, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                mArrayIn = Array(Rec!intVoucherID, gbTransactionDate)
                objdb.ExecuteSP "spSaveReverseVouchers", mArrayIn, mArrayOut, , mCnn
                If Rec!tnyVoucherTypeID = 20 Then
                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                End If
                Rec.MoveNext
            Wend
        Else
            MsgBox "No Journal to Reverse", vbInformation
        End If
        Rec.Close
        '-------------------------------------------------------------------------------------'

        mSQL = "Update faCancelledVouchers Set tnyApproveStatus = 2,dtCancellationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',intApproverID = " & gbUserID & " Where numReceiptNo = " & mPaymentOrderNo
        mCnn.Execute mSQL
        mCnn.Execute "Update faPayOrder Set tnyCancelled = 1 Where vchPayorderNo = " & mPaymentOrderNo
    End Sub
