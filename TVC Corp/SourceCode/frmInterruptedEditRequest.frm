VERSION 5.00
Begin VB.Form frmInterruptedEditRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interrupted Receipt Edit Request"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbReason 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   645
      Width           =   2505
   End
   Begin VB.CommandButton cmdSearchVouchers 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3855
      TabIndex        =   5
      Top             =   270
      Width           =   330
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   570
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1005
      Width           =   2460
   End
   Begin VB.TextBox txtReceiptNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2460
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Edit Request"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1118
      TabIndex        =   2
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   555
      TabIndex        =   7
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label lblReason 
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
      Left            =   690
      TabIndex        =   4
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lblReceiptNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt No"
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
      Left            =   390
      TabIndex        =   3
      Top             =   285
      Width           =   945
   End
End
Attribute VB_Name = "frmInterruptedEditRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub cmdSearchVouchers_Click()
        Dim mCnn  As New ADODB.Connection
        Dim Rec   As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql  As String
        
        '*********************************************************************************************'
        '                       Procedure to Search Interrupt Receipt Vouchers                        '
        '*********************************************************************************************'
        On Error GoTo Err
        frmSearchVouchers.CheckMode = 10
        frmSearchVouchers.chkPayment.Enabled = False
        frmSearchVouchers.chkContra.Enabled = False
        frmSearchVouchers.chkJournal.Enabled = False
        frmSearchVouchers.chkInterrupted.Visible = True
        frmSearchVouchers.chkInterrupted.value = 1
        
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            
                mSql = " SELECT  tnyVoucherGroupID,intVoucherID From faVouchers Where intVoucherID = " & gbSearchID & " "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If (IsNull(Rec!tnyVoucherGroupID)) Or Rec!tnyVoucherGroupID <> 4 Then 'Checking whether is it an Interrupt Receipt Or not
                        'If Rec!tnyVoucherGroupID <> 4 Then
                        MsgBox "It is not an Interrupt Receipt", vbInformation
                        Exit Sub
                    Else
                        txtReceiptNo.Text = gbSearchCode
                        txtReceiptNo.Tag = gbSearchID
                    End If
                    'End If
                End If
                Rec.Close
              Else
                MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
            End If
            gbSearchCode = ""
            gbSearchID = -1
      End If
      Exit Sub
Err:
      MsgBox Err.Description
    End Sub

    Private Sub cmdSend_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mAryIn      As Variant
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mVoucherNo  As String
        '*********************************************************************************************'
        '               Procedure to send request for Interrupt Receipt Cancellation                  '
        '*********************************************************************************************'
        On Error GoTo Err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If txtReceiptNo.Text = "" Then
            MsgBox "Please enter the Receipt No", vbInformation
            txtReceiptNo.SetFocus
            Exit Sub
        End If
        If txtReceiptNo.Tag = "" Then
            MsgBox "Please enter the Receipt No", vbInformation
            txtReceiptNo.SetFocus
            Exit Sub
        End If
        If cmbReason.ListIndex < 1 Then
            MsgBox "Please select the Reason", vbInformation
            cmbReason.SetFocus
            Exit Sub
        End If
        mVoucherNo = Token(txtReceiptNo.Text, "-")
        mSql = "Select * From faInterruptedRequests"
        mSql = mSql + " Where intVoucherID = " & txtReceiptNo.Tag
        mSql = mSql + " And tnyStatus <>0"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            MsgBox "There is already an Edit Request sent for this Receipt", vbInformation
            Exit Sub
        End If
        Rec.Close
        
        mAryIn = Array(gbCounterID, _
                       gbUserID, _
                       1, _
                       Date, _
                       3, _
                       cmbReason.ItemData(cmbReason.ListIndex), _
                       txtReason.Text, _
                       mVoucherNo, _
                       val(txtReceiptNo.Tag))
        'objdb.ExecuteSP "spSaveInterruptedRequest", mAryIn, , , mCnn, adCmdStoredProc'NOTE: NOT IN USE NOW- SP CHANGED
        MsgBox "Request sent to Nodal Officer", vbInformation
        cmdSend.Enabled = False
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub Form_Load()
        Dim mSql As String
        
        mSql = "Select vchCancelReason,intCancelID From faCancelReason WHERE intCancelID < 5"
        PopulateList cmbReason, mSql, , True, True, True, enuSourceString.Saankhya
    End Sub

    Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
