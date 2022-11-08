VERSION 5.00
Begin VB.Form frmManualReconcile 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reconcile "
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtVoucherNo 
      Height          =   285
      Left            =   2745
      TabIndex        =   3
      Top             =   525
      Width           =   1575
   End
   Begin VB.CommandButton cmdReconcile 
      BackColor       =   &H00DAF2F2&
      Caption         =   "&Reconcile"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label lblVoucher 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No (if Any)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1050
      TabIndex        =   1
      Top             =   510
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   0
      Top             =   990
      Width           =   675
   End
End
Attribute VB_Name = "frmManualReconcile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarVoucherFlag As Boolean

Private Sub cmdReconcile_Click()
    Dim mTokenID As Variant
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mLoopCount As Integer
    Dim mVoucherID As Double
    Dim mMultipleVouchers As Integer
    Dim mLoop As Integer
    Dim mFirstRowUpdate As Boolean
    
    mTokenID = val(txtVoucherNo.Tag)
    If mTokenID <> 0 Then
        On Error GoTo ErrorHandler:
        objDB.SetConnection mCnn
        If mvarVoucherFlag Then
            'Note:- Normal Reconciliation Mode
            frmBankReconcilationProcess.Remarks = Trim(txtRemarks.Text)
            Call frmBankReconcilationProcess.ReconcileVouchers
            Unload Me
        Else
            'Note:-Manual Reconciliation Mode
            With frmBankReconcilationProcess
                mFirstRowUpdate = False
                For mLoop = 1 To .fgBankStatement.Rows - 1
                    If .fgBankStatement.RowHidden(mLoop) = False And _
                        .fgBankStatement.Cell(flexcpChecked, mLoop, 5) = 2 And _
                            .fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked Then
                        mSql = "Update faBankReconciliationEntries Set vchRemarks = '" & Trim(txtRemarks) & "' ,"
                        mSql = mSql + " intVoucherNo =  " & val(txtVoucherNo)
                        mSql = mSql + ", numTockenID =  " & mTokenID
                        mSql = mSql + ", tnyReconciled = 3"
                        mSql = mSql + ", dtReconcileDate = getDate()"
                        If mFirstRowUpdate Then
                            mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID), 1) From faBankReconciliationEntries A )"
                        Else
                            mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID)+1, 1) From faBankReconciliationEntries A )"
                        End If
                        mSql = mSql + " Where intReconciliationID = " & .fgBankStatement.TextMatrix(mLoop, 0)
                        mCnn.Execute mSql
                        mFirstRowUpdate = True
                        frmBankReconcilationProcess.ManuallyReconciledFlag = True
                        frmBankReconcilationProcess.vsTitleGrid.Clear 1, 1
                        frmBankReconcilationProcess.vsTitleGrid.Tag = ""
                    End If
                Next mLoop
            End With
        End If
        Unload Me
    Else
        mSql = "Didn't able to identify the Bank Scroll Entry" & vbCrLf
        mSql = mSql + " Please Double Click and select the Entry to Reconcile Manually"
        MsgBox mSql, vbInformation
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Didn't able Reconcile the Entry Manually, Please try again :" & Error$, vbInformation
    Call Form_Load
End Sub

Private Sub Form_Load()
    txtVoucherNo.Text = ""
    txtVoucherNo.Tag = ""
    txtRemarks.Text = ""
    If mvarVoucherFlag Then
        lblVoucher.Visible = False
        txtVoucherNo.Visible = False
    Else
        lblVoucher.Visible = True
        txtVoucherNo.Visible = True
    End If
End Sub

Public Property Let VoucherFlag(ByVal mData As Boolean)
    'Note:- To Identify Whether its Normal Reconciliation or Manual Reconciliation
    '       Normal Reconciliation >> mDate = True
    '       Manual Reconciliation >> mData = False
    mvarVoucherFlag = mData
End Property
