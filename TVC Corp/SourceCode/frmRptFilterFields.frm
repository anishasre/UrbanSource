VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmRptFilterFields 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Fields"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00A0CBDF&
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2348
      TabIndex        =   17
      Top             =   3810
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00A0CBDF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3308
      TabIndex        =   16
      Top             =   3810
      Width           =   885
   End
   Begin VB.Frame fmeAccountHead 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Account Head"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   495
      TabIndex        =   10
      Top             =   1215
      Width           =   5925
      Begin VB.CommandButton cmdAccountHeadSearch 
         BackColor       =   &H00A0CBDF&
         Caption         =   "..."
         Height          =   300
         Left            =   3465
         TabIndex        =   13
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox txtAccountHead 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1665
         TabIndex        =   12
         Top             =   690
         Width           =   3870
      End
      Begin VB.TextBox txtAccountHeadCode 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   1665
         TabIndex        =   11
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   570
         TabIndex        =   15
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account HeadCode"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   390
         Width           =   1410
      End
   End
   Begin VB.TextBox txtFund 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1530
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   420
      Width           =   3585
   End
   Begin VB.CommandButton cmdSearchFund 
      BackColor       =   &H00A0CBDF&
      Caption         =   "..."
      Height          =   300
      Left            =   5130
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
   End
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Text            =   "01-Jan- 2008"
      Top             =   840
      Width           =   1515
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Top             =   840
      Width           =   1515
   End
   Begin VB.Frame fmeSubLedger 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000002&
      Height          =   1155
      Left            =   480
      TabIndex        =   0
      Top             =   2550
      Width           =   5925
      Begin VB.CommandButton cmdSearchSubledger 
         BackColor       =   &H00A0CBDF&
         Caption         =   "..."
         Height          =   300
         Left            =   5565
         TabIndex        =   4
         Top             =   270
         Width           =   300
      End
      Begin VB.TextBox txtSubLedger 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1080
         TabIndex        =   3
         Top             =   645
         Width           =   4455
      End
      Begin VB.ComboBox cmbSubLedgerType 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtSubLedgerCode 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   3720
         TabIndex        =   1
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Ledger"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   285
         Width           =   825
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   8400
      Top             =   3720
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   360
      Left            =   3060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65601537
      CurrentDate     =   39612
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   360
      Left            =   5130
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   840
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65601537
      CurrentDate     =   39612
   End
   Begin VB.Label lblTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   270
      TabIndex        =   23
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   1125
      TabIndex        =   22
      Top             =   420
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&To"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   3375
      TabIndex        =   21
      Top             =   900
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&From"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   1125
      TabIndex        =   20
      Top             =   885
      Width           =   375
   End
End
Attribute VB_Name = "frmRptFilterFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    
        Private mrptId As Integer
        Public mSelect As Boolean
' function for, from which Report menu is selected this form below :- Sinoj
    
    Public Property Let rptNames(rptMnuId As Integer)
            mrptId = rptMnuId
    End Property
    
    Private Sub cmdAccountHeadSearch_Click()
        Dim mSql As String
        If mrptId = 8 Then
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads"
        Else
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faBank
        End If
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        txtAccountHeadCode.SetFocus
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

''    Private Sub cmdSearchSubLedger_Click()
''        mSelect = True
''        frmSearchSubsidiaryAccount.Show 1
''    End Sub

    Private Sub cmdShow_Click()
'        Dim mRptFileName As String
'        Dim mInputPara As Variant
'        Dim mRpt As New CRAXDRT.Report
'        Dim mApp As New CRAXDRT.Application
'        Dim objCrv As CRViewer9
'
'        mRptFileName = App.Path & "\Reports\rptLedgerTrialBalance.rpt"
'        Set mRpt = Nothing
'        mApp.LogOnServer "ODBC", "dsnFA", "DB_Finance", "FAUser", "FAUser"
'        Set mRpt = mApp.OpenReport(mRptFileName, 1)
'        objCrv.ReportSource = mRpt
'        objCrv.ViewReport

        Dim mLoop As Integer
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
'        Dim mVoucherType    As String
'        mVoucherType = IIf(chkR.Value = 1, chkR.Tag, "") + IIf(chkP.Value = 1, chkP.Tag, "") + IIf(chkC.Value = 1, chkC.Tag, "") + IIf(chkJ.Value = 1, chkJ.Tag, "")
        Select Case mrptId
        
            Case 1
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), val(txtFund.Tag))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptLedgerTrialBalance.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 2
                
                arInput = Array(CDate(txtToDate.Text))
                frmNewViewer.rptFileName = App.Path & "\Reports\rptB1Schedule.rpt"
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.InputParameters = arInput
                Call frmNewViewer.ShowReport
                Unload Me
                frmNewViewer.Show
                Dim arrInput As Variant
                arrInput = Array(arInput(0), val(txtFund))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptBalanceSheetSchedule.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arrInput
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
                Dim frmNewViewer1 As New frmRptViewer
                'arInput = Array(CDate(txtToDate.Text))
                frmNewViewer1.rptFileName = App.Path & "\Reports\rptBalanceSheet.rpt"
                frmNewViewer1.WindowState = vbMaximized
                frmNewViewer1.WindowState = vbMaximized
                frmNewViewer1.InputParameters = arrInput
                Call frmNewViewer1.ShowReport
                frmNewViewer1.Show
            Case 3
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), val(txtFund.Tag))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptIESchedules.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Call frmNewRpt.ShowReport
                Unload Me
                frmNewRpt.Show
                
                'arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewViewer.rptFileName = App.Path & "\Reports\rptIncomeAndExpenditure.rpt"
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.InputParameters = arInput
                Call frmNewViewer.ShowReport
                frmNewViewer.Show
            Case 4
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), val(txtFund.Tag))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptReceiptPaymentSchedules.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Call frmNewRpt.ShowReport
                Unload Me
                frmNewRpt.Show
                
                'arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewViewer.rptFileName = App.Path & "\Reports\rptReceiptsPayments.rpt"
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.InputParameters = arInput
                Call frmNewViewer.ShowReport
                frmNewViewer.Show
            Case 5
                arInput = Array(val(txtAccountHeadCode.Tag), CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptCashBook.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 6
                
                arInput = Array(val(txtAccountHeadCode.Tag), CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptBankBook.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 7
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), gbFundID)
                frmNewRpt.rptFileName = App.Path & "\Reports\rptJournalBook.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 8
                
                arInput = Array(val(txtAccountHeadCode.Tag), CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGeneralLedger.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 9
                frmNewRpt.rptFileName = App.Path & "\Reports\rptAppropriationControl.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 10
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-36AssetReplacementRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 11
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN41AuthorisationIssuetoSecretarys.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 12
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-22BillofReceipts.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 13
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-11CollectionRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 14
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-21DemandRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 15
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-19DepositReceivedRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 16
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-30DocumentControl.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 17
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-40.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 18
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-35FunctionWiseExpenditureSubsidiaryLedger.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 19
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-34FunctionWiseReceiptSubsidiaryLedger.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 20
                frmNewRpt.rptFileName = App.Path & "\Reports\RegisterOfFundsReceved.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 21
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-31RegisterOfImmovableProperty.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 22
                frmNewRpt.rptFileName = App.Path & "\Reports\rptImplementingOfficerWiseAllotment.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 23
                frmNewRpt.rptFileName = App.Path & "\Reports\rptRegisterOfIncomeAndExpenditure.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 24
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-33RegisterOfLand.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 25
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-42LetterOfAllotment.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 26
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-12MemorandumofCollection.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 27
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-32RegisterOfMovableProperty.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 28
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-8OfficialReceipt.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 29
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-15PaymentOrder.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 30
                frmNewRpt.rptFileName = App.Path & "\Reports\rptProjectRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 31
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-17RegisterOfAdvances.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 32
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-14RegisterofBillsforPayment"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 33
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-18RegisterOfPermanentAdvance.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 34
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-37RegisterOfPublicLightingSystem.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 35
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-38RequesitionforReleaseOfFundCodes.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 36
                frmNewRpt.rptFileName = App.Path & "\Reports\rptStatementOfOutstandingLiabilityForExpenses.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 37
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-10StatementOnStatusofChequesReceived.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 38
                frmNewRpt.rptFileName = App.Path & "\Reports\rptSubsidaryRegister.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 39
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-13SummaryofCollection.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 40
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-23SummaryStatementsofBills.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 41
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-20SummaryStatementOfDeposits.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 42
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-27SummaryStatementOfRefundAndRemission.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 43
                frmNewRpt.rptFileName = App.Path & "\Reports\rptGEN-28SummaryStatementOfWriteOffs.rpt"
                frmNewRpt.WindowState = vbMaximized
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 44
                arInput = Array(CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptBudgeVariation.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
             Case 45
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), val(txtFund.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptChequeIssueRegisterGEN16.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 46
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), val(txtFund.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptChequeReceivedRegisterGEN9.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            Case 47
                arInput = Array(val(txtAccountHeadCode.Tag), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptBankReconciliationDetails.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Call frmNewRpt.ShowReport
                Unload Me
                frmNewRpt.Show
                
                'arInput = Array(Val(txtAccountHeadCode.Tag), CDate(txtToDate.Text))
                frmNewViewer.rptFileName = App.Path & "\Reports\rptBankReconciliation.rpt"
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.InputParameters = arInput
                Call frmNewViewer.ShowReport
                frmNewViewer.Show
                
            Case 48
                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptCashFlow.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
                
            Case 49
                arInput = Array(CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptKeyRatios.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
            
        End Select
    End Sub
    
    Private Sub FormInitilaize()
        Dim objFund As New clsFund
        objFund.SetFund (gbFundID)
        If objFund.FundID > -1 Then
            txtFund.Text = objFund.FundName
            txtFund.Tag = objFund.FundID
        Else
            txtFund.Text = ""
            txtFund.Tag = ""
        End If
        txtFromDate.Text = DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbEndingDate)
        dtpFrom.Value = gbStartingDate
        dtpToDate.Value = gbEndingDate
        txtAccountHeadCode.Text = ""
        txtAccountHeadCode.Tag = ""
        txtAccountHead.Text = ""
        cmbSubLedgerType.ListIndex = -1
        txtSubLedgerCode.Text = ""
        txtSubLedger.Text = ""
        mSelect = False
        Select Case mrptId
            Case 5: txtAccountHeadCode.Tag = gbAcHeadIDCash
                    txtAccountHeadCode.Text = gbAcHeadCodeCash
                    txtAccountHead.Text = "Cash"
                    fmeAccountHead.Enabled = False
                    fmeSubLedger.Enabled = False
            Case 6: txtSubLedger.Enabled = False
                    txtSubLedgerCode.Enabled = False
                    cmbSubLedgerType.Enabled = False
                    cmdSearchSubLedger.Enabled = False
            Case 7: fmeAccountHead.Enabled = False
            Case 8: txtSubLedger.Enabled = False
                    txtSubLedgerCode.Enabled = False
                    cmbSubLedgerType.Enabled = False
                    cmdSearchSubLedger.Enabled = False
            Case 1, 3, 4, 48: cmdAccountHeadSearch.Enabled = False
                          txtAccountHead.Enabled = False
                          txtAccountHeadCode.Enabled = False
            Case 2, 49: cmdAccountHeadSearch.Enabled = False
                    txtAccountHead.Enabled = False
                    txtAccountHeadCode.Enabled = False
                    txtFromDate.Enabled = False
                    Label4.Caption = "&On"
            Case 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, _
            19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, _
            30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43: fmeAccountHead.Enabled = False
                    fmeSubLedger.Enabled = False
                    txtFromDate.Enabled = False
                    txtToDate.Enabled = False
            Case 45, 46
                    fmeSubLedger.Enabled = False
                    fmeAccountHead.Enabled = False
            Case 47
                    txtFromDate.Visible = False
                    dtpFrom.Visible = False
                    Label3.Visible = False
                    Label4.Left = 2200
                    Label4.Caption = "Date"
                    txtToDate.Left = 2580
                    dtpToDate.Left = 4110
                    txtAccountHeadCode.Tag = 1506
                    txtAccountHeadCode.Text = "450210100"
                    txtAccountHead.Text = "SBT A/C -Own Fund"
                    
        End Select
    End Sub
    
    Private Sub dtpFrom_CloseUp()
        txtFromDate.Text = DdMmmYy(dtpFrom.Value)
    End Sub
    
    Private Sub dtpFrom_DropDown()
        If IsDate(txtFromDate) Then
            dtpFrom.Value = txtFromDate.Text
        Else
            dtpFrom.Value = gbTransactionDate
        End If
    End Sub
    
    Private Sub dtpToDate_CloseUp()
        txtToDate.Text = DdMmmYy(dtpToDate.Value)
    End Sub
    
    Private Sub dtpToDate_DropDown()
        If IsDate(txtToDate) Then
            dtpToDate.Value = txtToDate.Text
        Else
            dtpToDate.Value = gbTransactionDate
        End If
    End Sub
    
    Private Sub Form_Load()
'        If CheckTransactionsCorrect Then
'            lblTrans.Caption = "All Transactions are Correct"
'        Else
'            lblTrans.Caption = "Some Transaction/Transactions is/are not matching"
'        End If
        Call FormInitilaize
        txtToDate.Text = DdMmmYy(gbTransactionDate)
    End Sub
    
    Private Sub txtAccountHeadCode_GotFocus()
        If Len(gbSearchStr) Then
            Dim objAccHead As New clsAccounts
            objAccHead.SetAccountCode (Token(gbSearchStr, " "))
            If objAccHead.AccountHeadID > 0 Then
                txtAccountHeadCode.Text = objAccHead.AccountCode
                txtAccountHeadCode.Tag = objAccHead.AccountHeadID
                txtAccountHead = objAccHead.AccountHead
            End If
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub
    
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
    End Sub

    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub

    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
            MsgBox "Please Enter a valid Date", vbInformation
            txtFromDate.Text = ""
            txtFromDate.SetFocus
            Exit Sub
        End If
    End Sub

    Private Function CheckTransactionsCorrect() As Boolean
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim intFile As Integer
        intFile = FreeFile
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connection to Saankhya not present" & vbNewLine & "Contact System Administrator"
            Exit Function
        End If
        mSql = "Select C.intTransactionID, Sum(First1)[FSum],Sum(Second1)[SSum],dtTransactionDate,intVoucherID,GetDate()[CurDate] From( Select " & _
        "A.intTransactionID,Case When intByAccountHeadID is Null Then  " & _
        "Sum(fltAmount) " & _
        "Else " & _
        "0 " & _
        "End [First1], " & _
        "Case When intByAccountHeadID is not Null Then " & _
        "Sum (fltAmount) " & _
        "Else " & _
        "0 " & _
        "End [Second1],dtTransactionDate,intVoucherID " & _
        "From fatransactions A " & _
        "Inner Join faTransactionChild B On A.intTransactionID = B.intTransactionID  " & _
        "Group By A.intTransactionID,intByAccountHeadID,dtTransactionDate,intVoucherID)C " & _
        "Group By C.intTransactionID,dtTransactionDate,intVoucherID Having intTransactionID <> 0 And (Sum(First1)- Sum(Second1)) <> 0"

        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Open "C:\txt1.txt" For Output As intFile
            While Not Rec.EOF
                Print #intFile, "Transaction ID " & Rec!intTransactionID
                Print #intFile, "Sum Mismatch 1 " & Rec!FSum
                Print #intFile, "Sum Mismatch 2 " & Rec!SSum
                Print #intFile, "TransactinDate " & Rec!dtTransactionDate
                Print #intFile, "VoucherID " & Rec!intVoucherID
                Print #intFile, "Current Date " & Rec!CurDate
                Print #intFile, "---------------------------------------------------------------------" & vbNewLine
                CheckTransactionsCorrect = False
                Rec.MoveNext
            Wend
            Close
        Else
            CheckTransactionsCorrect = True
        End If
    End Function
