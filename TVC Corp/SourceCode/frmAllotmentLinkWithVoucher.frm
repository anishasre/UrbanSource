VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAllotmentLinkWithVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AllotmentLinkWithVoucher"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   360
      Left            =   5640
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdApprove 
      Appearance      =   0  'Flat
      Caption         =   "Approve"
      Height          =   360
      Left            =   4560
      TabIndex        =   39
      Top             =   4920
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Voucher Details"
      Height          =   3615
      Left            =   5475
      TabIndex        =   8
      Top             =   525
      Width           =   5295
      Begin VB.TextBox txtVoucherAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   31
         Top             =   2790
         Width           =   1755
      End
      Begin VB.TextBox txtVoucherTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   30
         Top             =   2415
         Width           =   2490
      End
      Begin VB.TextBox txtBank 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Top             =   2010
         Width           =   2490
      End
      Begin VB.TextBox txtInstrumentNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   1650
         Width           =   1710
      End
      Begin VB.TextBox TxtInstrumentType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   1260
         Width           =   3030
      End
      Begin VB.TextBox txtVoucherDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   855
         Width           =   3030
      End
      Begin VB.TextBox txtVoucherNoList 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         TabIndex        =   25
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label18 
         Caption         =   "Amount"
         Height          =   255
         Left            =   1005
         TabIndex        =   38
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label Label17 
         Caption         =   "Transaction Type"
         Height          =   285
         Left            =   345
         TabIndex        =   37
         Top             =   2430
         Width           =   1350
      End
      Begin VB.Label Label16 
         Caption         =   "Bank"
         Height          =   225
         Left            =   1095
         TabIndex        =   36
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label15 
         Caption         =   "Instrument No:"
         Height          =   210
         Left            =   555
         TabIndex        =   35
         Top             =   1695
         Width           =   1155
      End
      Begin VB.Label Label14 
         Caption         =   "Instrument Type"
         Height          =   210
         Left            =   360
         TabIndex        =   34
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label13 
         Caption         =   "Voucher Date"
         Height          =   210
         Left            =   480
         TabIndex        =   33
         Top             =   855
         Width           =   1140
      End
      Begin VB.Label Label12 
         Caption         =   "Voucher No:"
         Height          =   240
         Left            =   600
         TabIndex        =   32
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details of LetterOfAuthority"
      Height          =   3615
      Left            =   105
      TabIndex        =   7
      Top             =   510
      Width           =   5295
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   3120
         Width           =   3345
      End
      Begin VB.TextBox txtFunctionaries 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   2700
         Width           =   3345
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   2310
         Width           =   3345
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1935
         Width           =   1440
      End
      Begin VB.TextBox TxtCategory 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   1560
         Width           =   3345
      End
      Begin VB.TextBox TxtSource 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   3345
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   3345
      End
      Begin VB.TextBox txtAllotmentNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label11 
         Caption         =   "Account Head"
         Height          =   285
         Left            =   480
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Functionaries  "
         Height          =   180
         Left            =   600
         TabIndex        =   23
         Top             =   2700
         Width           =   1020
      End
      Begin VB.Label Label9 
         Caption         =   "Function  "
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   2310
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Amount "
         Height          =   240
         Left            =   960
         TabIndex        =   21
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label Label6 
         Caption         =   "Category "
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Source   "
         Height          =   210
         Left            =   960
         TabIndex        =   19
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Transaction Type"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Allotment No  "
         Height          =   210
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10860
      TabIndex        =   6
      Top             =   0
      Width           =   10860
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   -3450
      Top             =   6195
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdVerify 
      Appearance      =   0  'Flat
      Caption         =   "Verify"
      Height          =   360
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearchVoucher 
      Caption         =   "..."
      Height          =   300
      Left            =   7755
      TabIndex        =   3
      Top             =   4455
      Width           =   315
   End
   Begin VB.TextBox txtVoucherNo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5700
      TabIndex        =   2
      Top             =   4455
      Width           =   2040
   End
   Begin VB.TextBox txtDemandNo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2580
      TabIndex        =   1
      Top             =   4455
      Width           =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher No:"
      Height          =   285
      Left            =   4740
      TabIndex        =   4
      Top             =   4455
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Demand No:"
      Height          =   300
      Left            =   1500
      TabIndex        =   0
      Top             =   4455
      Width           =   1005
   End
End
Attribute VB_Name = "frmAllotmentLinkWithVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmdApprove_Click()
    Dim mSql    As String
    Dim objDb   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    
        mSql = "Update faAllotmentRegister set tnyStatus=5 where vchAllotmentNo=" & txtAllotmentNo.Text & "  "
        objDb.ExecuteSP mSql, , , , mCnn, adCmdText
        cmdApprove.Enabled = False
'        cmdReject.Enabled = False
    
    End Sub

'''    Private Sub cmdReject_Click()
'''        frmReject.Mode = 12
'''        frmReject.RequestTypeID = txtAllotmentNo.Text
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''        cmdApprove.Enabled = False
'''    End Sub

    Private Sub cmdSearchVoucher_Click()
    Dim mSql    As String
    Dim objDb   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
        frmSearchVouchers.CheckMode = 10
        frmSearchVouchers.chkContra.Visible = False
        frmSearchVouchers.chkReceipt.value = 1
        frmSearchVouchers.chkJournal.Visible = False
        frmSearchVouchers.chkPayment.Visible = False
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            txtVoucherNo.Text = gbSearchCode
            txtVoucherNo.Tag = gbSearchID
            gbSearchCode = ""
            gbSearchID = -1
        End If
        If objDb.SetConnection(mCnn) Then
            If txtVoucherNo.Tag <> "" Then
                mSql = " SELECT intVoucherID, intVoucherNo, fltAmount From faVouchers"
                mSql = mSql + " Where intVoucherID = " & txtVoucherNo.Tag & " "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If Rec!fltAmount <> txtAmount.Text Then
                        MsgBox "Please Enter the Correct Voucher.Amount incorrect", vbInformation, "Saankhya"
                        txtVoucherNo.Text = ""
                        txtVoucherNo.Tag = ""
                        txtDemandNo.Text = ""
                        Exit Sub
                    End If
                End If
                Rec.Close
            End If
        End If
        If objDb.SetConnection(mCnn) Then
            If txtVoucherNo.Tag <> "" Then
                mSql = "SELECT faVouchers.intVoucherID, faVouchers.intVoucherNo, faIDemandTBL.numDemandID, faIDemandTBL.vchDemandNo, faVouchers.fltAmount"
                mSql = mSql + " FROM faVouchers INNER JOIN"
                mSql = mSql + " faIDemandTBL ON faVouchers.intVoucherID = faIDemandTBL.intVoucherID AND faVouchers.intLocalBodyID = faIDemandTBL.intLBID"
                mSql = mSql + " Where  faVouchers.intVoucherID = " & txtVoucherNo.Tag & " "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If Rec!vchDemandNo <> "" Then
                         txtDemandNo.Text = (Rec!vchDemandNo)
                    End If
                End If
                Rec.Close
            End If
        End If
        Call FnVerify
    End Sub

    Private Function FnVerify() As Boolean   'Private Sub cmdVerify_Click() CHANGED
        Dim mCnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim mSql    As String
        Dim Rec As New ADODB.Recordset
        
        If txtVoucherNo.Tag <> "" Then
            mSql = "Update faAllotmentRegister set tnyStatus=4, intVoucherID= " & txtVoucherNo.Tag & " where vchAllotmentNo=" & txtAllotmentNo.Text & "  "
            objDb.ExecuteSP mSql, , , , mCnn, adCmdText
            cmdVerify.Enabled = False
            If objDb.SetConnection(mCnn) Then
                mSql = "SELECT faVouchers.intVoucherID, faTransactionType.vchTransactionType, faVouchers.intVoucherNo, faVouchers.vchInstrumentNo, faVouchers.fltAmount, "
                mSql = mSql + " faVouchers.vchBank , faInstrumentTypes.vchInstrumentType, faVouchers.dtDate FROM faVouchers INNER JOIN"
                mSql = mSql + " faTransactionType ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                mSql = mSql + " faInstrumentTypes ON faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
                mSql = mSql + " Where faVouchers.intVoucherID = " & txtVoucherNo.Tag & " "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    txtVoucherNoList.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    txtVoucherDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    TxtInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                    txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                    txtVoucherTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    txtVoucherAmount.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                End If
                Rec.Close
            End If
            MsgBox "Voucher Linked Successfully", vbInformation, "Saankhya"
        End If
    End Function


    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FormInitialize
        Frame2.Enabled = False
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdApprove.Visible = False
'            cmdReject.Visible = False
            cmdVerify.Visible = False
        End If
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then     'Accounts Officer   gbUserTypeID = 4 Then
            cmdApprove.Visible = True
'            cmdReject.Visible = True
            cmdVerify.Enabled = False
            cmdSearchVoucher.Enabled = False
            txtDemandNo.Enabled = False
            txtVoucherNo.Enabled = False
        End If
    End Sub
    Private Sub FormInitialize()
        txtDemandNo.Text = ""
        txtVoucherNo.Text = ""
        
    End Sub
    Private Sub Form_Activate()
        Me.Top = 2500
        Me.Left = (frmMenu.Width - Me.Width) / 2
     End Sub
