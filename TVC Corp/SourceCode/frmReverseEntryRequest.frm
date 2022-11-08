VERSION 5.00
Begin VB.Form frmReverseEntryRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Entry "
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "frmReverseEntryRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeReversedVoucherList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   7260
      TabIndex        =   41
      Top             =   1170
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton cmdRev 
         Caption         =   "¬"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   43
         Top             =   150
         Width           =   585
      End
      Begin VB.CommandButton cmdFwd 
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4050
         TabIndex        =   42
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Multiple Reversed Vouchers"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   660
         TabIndex        =   44
         Top             =   180
         Width           =   3375
      End
   End
   Begin VB.ListBox lstVouchers 
      Height          =   1035
      Left            =   9300
      TabIndex        =   40
      Top             =   4320
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Frame fmeChequeReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   60
      TabIndex        =   36
      Top             =   1140
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5670
         TabIndex        =   38
         Top             =   150
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¬"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   270
         TabIndex        =   37
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Multiple Vouchers For Cheqe Return"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   960
         TabIndex        =   39
         Top             =   180
         Width           =   4635
      End
   End
   Begin VB.TextBox txtReverseDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3255
      Width           =   2640
   End
   Begin VB.CommandButton cmdReversedTransaction 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11280
      TabIndex        =   20
      Top             =   2595
      Width           =   300
   End
   Begin VB.TextBox txtReverseVoucherNo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2580
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4740
      Left            =   6960
      TabIndex        =   16
      Top             =   870
      Width           =   45
   End
   Begin VB.CommandButton cmdSearchReason 
      Caption         =   "..."
      Height          =   315
      Left            =   6405
      TabIndex        =   13
      Top             =   2925
      Width           =   300
   End
   Begin VB.CommandButton cmdSeat 
      Caption         =   "..."
      Height          =   315
      Left            =   4185
      TabIndex        =   12
      Top             =   3915
      Width           =   300
   End
   Begin VB.TextBox txtReason 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2910
      Width           =   4845
   End
   Begin VB.CommandButton cmdSearchVoucher 
      Caption         =   "..."
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   2250
      Width           =   300
   End
   Begin VB.TextBox txtVoucherNo 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2220
      Width           =   2640
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   11970
      TabIndex        =   6
      Top             =   5760
      Width           =   12030
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
         Left            =   6011
         TabIndex        =   35
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   7811
         TabIndex        =   34
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton cmmVerify 
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
         Left            =   2494
         TabIndex        =   32
         Top             =   60
         Width           =   1725
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
         Left            =   4256
         TabIndex        =   29
         Top             =   60
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   11970
      TabIndex        =   5
      Top             =   0
      Width           =   12030
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   9750
         Picture         =   "frmReverseEntryRequest.frx":1CCA
         Top             =   -420
         Width           =   2550
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReverseEntryRequest.frx":2FFD
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1605
         TabIndex        =   31
         Top             =   195
         Width           =   7920
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverse Entry "
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
         Left            =   270
         TabIndex        =   30
         Top             =   90
         Width           =   1185
      End
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
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   3900
      Width           =   2640
   End
   Begin VB.TextBox txtRemarks 
      Height          =   615
      Left            =   1530
      TabIndex        =   3
      Top             =   3270
      Width           =   4830
   End
   Begin VB.Label lblMsgBox 
      Alignment       =   2  'Center
      Caption         =   "Message Box"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   0
      TabIndex        =   33
      Top             =   5340
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dated"
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
      Left            =   8055
      TabIndex        =   28
      Top             =   3270
      Width           =   510
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "User"
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
      Left            =   8160
      TabIndex        =   27
      Top             =   3630
      Width           =   390
   End
   Begin VB.Label lblReverseUser 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8610
      TabIndex        =   26
      Top             =   3585
      Width           =   2640
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7260
      TabIndex        =   24
      Top             =   3930
      Width           =   4650
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Reversed Voucher "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7260
      TabIndex        =   23
      Top             =   1920
      Width           =   4650
   End
   Begin VB.Label lblReverseAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8610
      TabIndex        =   22
      Top             =   2925
      Width           =   2640
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Net Amount"
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
      Left            =   7530
      TabIndex        =   21
      Top             =   2955
      Width           =   1035
   End
   Begin VB.Label lblReverseVoucherType 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8610
      TabIndex        =   19
      Top             =   2220
      Width           =   2640
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   7605
      TabIndex        =   17
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label lblNetAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1530
      TabIndex        =   15
      Top             =   2565
      Width           =   2640
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Net Amount"
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
      Left            =   450
      TabIndex        =   14
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   855
      TabIndex        =   10
      Top             =   2925
      Width           =   630
   End
   Begin VB.Label lblVoucherType 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1530
      TabIndex        =   8
      Top             =   1860
      Width           =   2640
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Forwarded Seat "
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
      Left            =   120
      TabIndex        =   2
      Top             =   3930
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   525
      TabIndex        =   1
      Top             =   2235
      Width           =   975
   End
   Begin VB.Label lblReturnDate 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   0
      Top             =   3270
      Width           =   765
   End
End
Attribute VB_Name = "frmReverseEntryRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
        Dim mvarCategoryID  As Integer
        Dim intInstrumentTypeID As Variant
        Dim intMultipleVouchers As Integer
        Private intVerify As Integer    '       0 = not verified;   1 = verified        '
        Private intUserType As Integer  '       0 = operator;       1 = approver        '
        Private intRequestID As Long
    Private Sub cmdChequeReturn_Click()
'        Call FormInitialize
        cmdRequest.Enabled = True
        frmChequeBounceRequest.Show vbModal
    End Sub

    Private Sub VoucherDetails(mVoucherID As Double)
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim arrOut      As Variant
        Dim mSQL        As String
        mSQL = "Select vchuserName,* From faVouchers  "
        mSQL = mSQL + " Left Join faUser On faUser.numuserID=faVouchers.intUserID Where intVoucherID=" & mVoucherID
        If objDB.SetConnection(mCnn) Then
            Rec.Open mSQL, mCnn
            If Not Rec.EOF Then
                txtReverseVoucherNo.Text = Rec!intVoucherNo
                txtReverseVoucherNo.Tag = Rec!intVoucherID
                lblReverseAmount.Caption = Rec!fltAmount
                txtReverseDate = Rec!dtDate
                lblReverseUser.Caption = IIf(IsNull(Rec!vchuserName), "", Rec!vchuserName)
                Select Case Rec!tnyVoucherTypeID
                        Case 10:
                            lblReverseVoucherType.Caption = "Receipt Voucher"
                        Case 20:
                            lblReverseVoucherType.Caption = "Payment Voucher"
                        Case 30:
                            lblReverseVoucherType.Caption = "Contra Voucher"
                        Case 40:
                            lblReverseVoucherType.Caption = "Journal Voucher"
                    End Select
            End If
        End If
    End Sub
    Private Sub cmdFwd_Click()
        If lstVouchers.ListIndex = lstVouchers.ListCount - 1 Then
            lstVouchers.ListIndex = 0
        Else
            lstVouchers.ListIndex = lstVouchers.ListIndex + 1
        End If
        txtReverseVoucherNo.Text = lstVouchers.Text
        Call VoucherDetails(val(txtReverseVoucherNo.Text))
    End Sub
Private Sub cmdApprove_Click()
        Dim objDB       As New clsDB
        Dim objReverse  As New clsReverseProcess
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim arrOut      As Variant
        Dim mCnt        As Integer
        Dim mFlag       As Boolean
        Dim mVoucher    As Double
        If ApproveValidation Then
            arrOut = objReverse.ReverseProcess(intRequestID, val(txtReason.Tag))
            If IsNull(arrOut) Then
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Reverse Entry Process Failed for Voucher No " & txtVoucherNo.Text
            Else
                For mCnt = 0 To UBound(arrOut) - 1
                    If UBound(arrOut) > 1 Then
                        mFlag = True
                        lstVouchers.AddItem arrOut(mCnt)
                        lstVouchers.ItemData(lstVouchers.NewIndex) = mCnt
                        lstVouchers.ListIndex = 0
                        'txtReverseVoucherNo.Tag = arrOut(mCnt)
                    Else
                         If lblVoucherType.Caption = "Receipt Voucher" Then
                            If (txtReason.Tag) > 2 Then
                                
                            End If
                         End If
    '                    lstVouchers.AddItem arrOut(mCnt)
    '                    lstVouchers.ItemData(lstVouchers.NewIndex) = mCnt
                    End If
                Next
                If mFlag Then
                    fmeReversedVoucherList.Visible = True
                    MsgBox "Multiple Vouchers Reversed : Please Verify"
                    
                Else
                    mVoucher = arrOut(0) 'lstVouchers.ItemData(lstVouchers.ListIndex)
                    Call VoucherDetails(mVoucher)
                End If
                cmdApprove.Enabled = False
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Reverse Entry Approved"
                objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                'mSQL = "Update faReverseEntry set  numApprovedUserID=" & gbUserID & ",dtApprovedDate='" & gbTransactionDate & "' Where intRequestID=" & intRequestID
                mSQL = "Update faReverseEntry set  numAuthorisedByAO=" & gbUserID & ",dtAuthorisationDateAO='" & gbTransactionDate & "' Where intRequestID=" & intRequestID
                mCnn.Execute mSQL
            End If
        End If
    End Sub
    Private Function ApproveValidation()
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim arrOut      As Variant
        Dim mCnt        As Integer
        Dim mFlag       As Boolean
        Dim mVoucher    As Double
        ApproveValidation = True
        mSQL = "Select * From faReverseEntry Inner Join faReverseEntryChild On faReverseEntry.intRequestID=faReverseEntryChild.intRequestID"
        mSQL = mSQL + " Inner Join faVouchers On faReverseEntryChild.intVoucherID=faVouchers.intVoucherID Where faReverseEntry.intRequestID= " & intRequestID
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
        If Not (Rec.EOF Or Rec.BOF) Then
            If Rec!intFinancialYearID <> gbFinancialYearID Then
                MsgBox "Previous Financial year Voucher" & vbNewLine & "Not Allowed to Approve", vbInformation
                ApproveValidation = False
            End If
        End If
        
    End Function
    
    Private Sub cmdCancel_Click()
        Unload Me
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
            lblMsgBox.Caption = "Please Select the Seat"
            cmdSeat.SetFocus
            RequsetValidation = False
            Exit Function
        End If
    
        RequsetValidation = True
    End Function

    Private Sub cmdRequest_Click()
        On Error GoTo err:
            Dim objDB       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim arrIn       As Variant
            Dim arrOut      As Variant
            Dim mRequestID  As Integer
            Dim mSQL        As String
            
            If RequsetValidation = False Then Exit Sub
            
            If objDB.SetConnection(mCnn) Then
                arrIn = Array(-1, _
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
                            0)
        
                objDB.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
                
                If Not IsNumeric(arrOut) Then
                    mRequestID = arrOut(0, 0)
                End If
                
                arrIn = ""
                arrIn = Array(mRequestID, val(txtVoucherNo.Tag))
                objDB.ExecuteSP "spSaveReverseEntryChild", arrIn, , , mCnn, adCmdStoredProc
                
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Reverse Entry requested to Higher Authority"
                
                cmdRequest.Enabled = False
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub cmdRev_Click()
        If lstVouchers.ListIndex = 0 Then
            lstVouchers.ListIndex = lstVouchers.ListCount - 1
        Else
            lstVouchers.ListIndex = lstVouchers.ListIndex - 1
        End If
        txtReverseVoucherNo.Text = lstVouchers.Text
        Call VoucherDetails(val(txtReverseVoucherNo.Text))
    End Sub

    Private Sub cmdReversedTransaction_Click()
        If txtReverseVoucherNo.Tag <> "" Then
            frmViewVoucher.MultipleVouchers = False
            frmViewVoucher.FormName = "frmReverseEntryRequest"
            frmViewVoucher.ArrayIn = Array(txtReverseVoucherNo.Tag)
            frmViewVoucher.Show vbModal
        End If
    End Sub

    Private Sub cmdSearchReason_Click()
        On Error GoTo err:
            If txtVoucherNo.Text = "" Then
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Please Select Voucher Before Giving Reason"
                Exit Sub
            End If
            lblMsgBox.Visible = False
            If intInstrumentTypeID = gbInstrumentCash Then
                frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <> 1 and intReasonID <> 2"
            Else
                frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <> 1"
            End If
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Show vbModal
            txtReason.Text = gbSearchStr
            txtReason.Tag = gbSearchID

            gbSearchID = -1
            gbSearchStr = ""
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearchVoucher_Click()
        Dim mSQL    As String
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        lblMsgBox.Visible = False
        frmSearchVouchers.Show vbModal
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        If gbSearchID <> -1 Then
            
            mSQL = "Select tnyVoucherTypeID,isNull(intInstrumentTypeID,100) intInstrumentTypeID From faVouchers Where intVoucherID=" & gbSearchID
            Rec.Open mSQL, mCnn
            If Not Rec.EOF Then
                If Rec!tnyVoucherTypeID = 20 Then
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Payment Voucher is Not Allowed To Reverse"
                    gbSearchCode = ""
                    gbSearchStr = ""
                    gbSearchID = -1
                    Exit Sub
                ElseIf Rec!tnyVoucherTypeID = 30 Then
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "Already sent Request for this Voucher"
                            cmmVerify.Enabled = False
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "This Voucher Already Reversed"
                            cmmVerify.Enabled = False
                        Else
                            Call GetVoucherDetails(gbSearchID)
                            cmmVerify.Enabled = True
                        End If
                ElseIf Rec!tnyVoucherTypeID = 40 Then
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "Already sent Request for this Voucher"
                            cmmVerify.Enabled = False
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "This Voucher Already Reversed"
                            cmmVerify.Enabled = False
                        Else
                            Call GetVoucherDetails(gbSearchID)
                            cmmVerify.Enabled = True
                        End If
                ElseIf Rec!tnyVoucherTypeID = 10 Then
                    If Rec!intInstrumentTypeID = 5 Then
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "Already sent Request for this Voucher"
                            cmmVerify.Enabled = False
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "This Voucher Already Reversed"
                            cmmVerify.Enabled = False
                        Else
                            Call GetVoucherDetails(gbSearchID)
                            cmmVerify.Enabled = True
                        End If
                    Else
                        lblMsgBox.Visible = True
                        lblMsgBox.Caption = "Receipt Voucher Other Than Cheque Instrument is Not Allowed To Reverse"
                        gbSearchCode = ""
                        gbSearchStr = ""
                        gbSearchID = -1
                        Exit Sub
                        End If
                End If
            End If
        End If
        gbSearchCode = ""
        gbSearchStr = ""
        gbSearchID = -1
    End Sub

    Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            If objDB.SetConnection(mCnn) Then
                mSQL = " Select tnyStatus from faReverseEntry "
                mSQL = mSQL + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSQL = mSQL + " Where intVoucherID =  " & VchID
                mSQL = mSQL + " And tnyStatus<>4"
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then      'Requested
                        CheckReverseRequestExist = 1
                    ElseIf Rec!tnyStatus = 1 Then  ' Approved
                        CheckReverseRequestExist = 2
                    Else                           'Cancelled Status=4
                        CheckReverseRequestExist = 3
                    End If
                    Exit Function
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact your System Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Public Function GetVoucherDetails(ByVal intVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mSQL As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim objDB As New clsDB
            
            If objDB.SetConnection(mCnn) Then
                mSQL = "Select * from faVouchers Where intVoucherID = " & intVoucherID
                Rec.Open mSQL, mCnn
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
                    lblNetAmount.Caption = Rec!fltAmount
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
    
    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtSeat.Text = gbSearchStr
            txtSeat.Tag = gbSearchID
        End If
    End Sub

'''    Private Sub cmdVoucherSearch_Click()
'''        Dim mVoucherID      As Double
'''        SelectionMode = 2
'''        Call FormInitialize
'''        cmdRequest.Enabled = True
'''        frmSearchTransactions.FormSelectionType = 2
'''        frmSearchTransactions.Show vbModal
'''        mVoucherID = gbSearchID
'''        frmSearchTransactions.FormSelectionType = -1
'''        Call FillSearchData(mVoucherID)
'''    End Sub

    
    Private Sub cmmVerify_Click()
        On Error GoTo err:

            
            If intMultipleVouchers > 1 Then
                frmViewVoucher.MultipleVouchers = True
                frmViewVoucher.ArrayIn = Array(RequestID)
            Else
            
                If txtVoucherNo.Tag = "" Then
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Please Select A Voucher to do Verification"
                    Exit Sub
                End If
            
                frmViewVoucher.MultipleVouchers = False
                frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
            End If
            
            If UserType = 1 Then
                frmViewVoucher.FormName = "frmReverseEntryRequest"
                frmViewVoucher.Show vbModal
                If VerifyStatus = 1 Then
                    cmdApprove.Enabled = True
                    cmmVerify.Enabled = False
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Please Click Request Button to Approve Reverse Entry Request"
                Else
                    cmdApprove.Enabled = False
                    cmmVerify.Enabled = True
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Voucher Verification Failed"
                End If
                
            Else
                frmViewVoucher.FormName = "frmReverseEntryRequest"
                frmViewVoucher.Show vbModal
                If VerifyStatus = 1 Then
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Please Click Request Button to Send Reverse Entry Request"
                    cmdRequest.Enabled = True
                    cmmVerify.Enabled = False
                Else
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Voucher Verification Failed"
                    cmdRequest.Enabled = False
                    cmmVerify.Enabled = True
                End If
            End If
            
            If intMultipleVouchers > 1 Then
                frmRequisition.Enabled = False
            Else
                frmRequisition.Enabled = True
            End If
            
            
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Load()
        On Error GoTo err:
            intMultipleVouchers = 0
            'If UserType = 1 Then
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                'cmdApprove.Enabled = True
                cmdRequest.Enabled = False
                Call FillData(RequestID)
                Call FormLock(True)
                cmdApprove.Enabled = True
            'ElseIf UserType = 0 Then
            ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                cmdApprove.Enabled = False
            'ElseIf UserType = 2 Then
            ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                cmdApprove.Enabled = False
                cmdRequest.Enabled = False
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Reverse Entry Process Completed for this request"
                Call FillData(RequestID)
            End If
            If intMultipleVouchers > 1 Then
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "There are Multiple Vouchers in this Request"
                cmdApprove.Enabled = False
                txtVoucherNo.Text = "Multiple Vouchers"
                lblNetAmount.Caption = Format(frmListReverseEntryRequests.vsGridForCheque.TextMatrix(frmListReverseEntryRequests.vsGrid.Row, 1), "0.00")
            End If
            
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub FormLock(mFlag As Boolean)
        'mflag =true for lock or disable
        cmdSearchVoucher.Enabled = Not mFlag
        txtVoucherNo.Locked = mFlag
        cmdSearchReason.Enabled = Not mFlag
        txtReason.Locked = mFlag
        txtRemarks.Locked = mFlag
        cmdSeat.Enabled = Not mFlag
        txtSeat.Locked = mFlag
    End Sub
    
    Private Function FillData(ByVal intRequestID As Long) As Boolean
         On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDB As New clsDB
            Dim mSQL As String
            
            If objDB.SetConnection(mCnn) Then
'                mSql = "Select * from faReverseEntry "
'                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID "
'                mSql = mSql + " Inner Join faReverseReasons On faReverseReasons.intReasonID = faReverseEntry.intReasonID "
'                mSql = mSql + " Left JOIN faSeats ON faReverseEntry.numForwardedSeatID = faSeats.numSeatID "
'                mSql = mSql + " Where faReverseEntry.intRequestID = " & intRequestID
                mSQL = "Select * from faReverseEntry "
                mSQL = mSQL + " Inner Join faReverseEntryChild On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID "
                mSQL = mSQL + " Inner Join faReasons On faReasons.intReasonID = faReverseEntry.intReasonID "
                mSQL = mSQL + " Left JOIN faSeats ON faReverseEntry.numForwardedSeatID = faSeats.numSeatID "
                mSQL = mSQL + " Where faReverseEntry.intRequestID = " & intRequestID
                
                Rec.Open mSQL, mCnn
                While Not (Rec.EOF Or Rec.BOF)
                    txtReason.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
                    txtReason.Tag = IIf(IsNull(Rec!intReasonID), "", Rec!intReasonID)
                    txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    txtSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                    txtSeat.Tag = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
                    
                    intMultipleVouchers = intMultipleVouchers + 1 '  To Count the No of Multiple Vouchers
                    Call GetVoucherDetails(Rec!intVoucherID)
                    Rec.MoveNext
                Wend
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function


    Private Function ReverEntry() As Boolean
        On Error GoTo err:
            Dim objDB As New clsDB
            Dim Rec As New ADODB.Recordset
        Exit Function
err:
        MsgBox (Error$)
    End Function


    Private Sub txtRemarks_LostFocus()
        lblMsgBox.Visible = False
    End Sub

    Private Sub txtSeat_KeyPress(KeyAscii As Integer)
       Call KeyPress(KeyAscii)
    End Sub
    Private Sub KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        Else
            KeyAscii = 0
        End If
    End Sub

    Public Property Let VerifyStatus(mData As Integer)
        intVerify = mData
    End Property
    
    Public Property Get VerifyStatus() As Integer
        VerifyStatus = intVerify
    End Property
    
    Public Property Let UserType(mData As Integer)
        intUserType = mData
    End Property
    
    Public Property Get UserType() As Integer
        UserType = intUserType
    End Property
    
    Public Property Let RequestID(mData As Integer)
        intRequestID = mData
    End Property
    
    Public Property Get RequestID() As Integer
        RequestID = intRequestID
    End Property


