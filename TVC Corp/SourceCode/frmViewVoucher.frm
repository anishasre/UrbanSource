VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewVoucher 
   BorderStyle     =   0  'None
   Caption         =   "frmViewVoucher"
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   12195
      TabIndex        =   1
      Top             =   6795
      Width           =   12255
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10380
         TabIndex        =   3
         Top             =   60
         Width           =   1725
      End
      Begin VB.CheckBox cmdVerify 
         Caption         =   "&Verify Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8610
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   6735
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12240
      lastProp        =   500
      _cx             =   21590
      _cy             =   11880
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmViewVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private aryStr  As String
    Private aryIn As Variant
    Private strFormName As String
    Private blnMultipleVouchers As Boolean

    Private Sub cmdClose_Click()
        frmReverseEntryRequest.VerifyStatus = 0
        frmReverseApproval.VerifyStatus = 0
        Unload Me
    End Sub

    Private Sub cmdPrint_Click()
        crvReport.PrintReport
        If FormName = "frmAFSClosingSourceOfFund" Then
            frmAFSClosingSourceOfFund.CheckPrintStatus = 1
            Unload Me
        
        End If
    End Sub

    Private Sub cmdVerify_Click()
        frmReverseEntryRequest.VerifyStatus = 1
        frmReverseApproval.VerifyStatus = 1
        frmReverseRequest.VerifyStatus = 1
        Unload Me
    End Sub

    Private Sub Form_Load()
        Call ReportView
        If FormName = "frmReverseEntryRequest" Or FormName = "frmListReverseEntryRequest" Then
            cmdVerify.Visible = True
        Else
            cmdVerify.Visible = False
        End If
    End Sub
    Private Sub ReportView()
         Dim Rpt As New CRAXDRT.Report
         Dim mApp As New CRAXDRT.Application
         Dim rptFileName As String
         Dim arrInput As Variant
         Dim mLoop As Long
         
         If FormName = "frmReverseEntryRequest" Then
            If MultipleVouchers Then
               rptFileName = App.Path & "\Reports\rptMultipleVoucher.rpt"
               crvReport.DisplayToolbar = True
               crvReport.EnableNavigationControls = True
               crvReport.EnableToolbar = True
               cmdVerify.Visible = True
            Else
               rptFileName = App.Path & "\Reports\rptVoucher.rpt"
            End If
         End If
         
         ''--------- Mofdified Reverse Entry--------------------------------
         If FormName = "frmReverseRequest" Then
            If MultipleVouchers Then
               rptFileName = App.Path & "\Reports\rptMultipleVoucher.rpt"
               crvReport.DisplayToolbar = True
               crvReport.EnableNavigationControls = True
               crvReport.EnableToolbar = True
            Else
               rptFileName = App.Path & "\Reports\rptVoucher.rpt"
               cmdVerify.Visible = True
            End If
         End If
         If FormName = "frmReverseDemand" Then
            rptFileName = App.Path & "\Reports\rptDemand.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableNavigationControls = True
            crvReport.EnableToolbar = True
            
            'ArrayIn = Array(CStr(ArrayIn))
         End If
         '------------------------------------------------------------------
         If FormName = "frmSubsidiaryCashBook" Then
            rptFileName = App.Path & "\Reports\rptVoucher.rpt"
         End If
         
          If FormName = "frmInterruptReceipt" Then
            rptFileName = App.Path & "\Reports\rptVoucher.rpt"
         End If
         
         If FormName = "frmViewPaymentOrder" Then
            rptFileName = App.Path & "\Reports\rptPOjournals.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableNavigationControls = True
            crvReport.EnableToolbar = True
            cmdPrint.Visible = True '************MODIFIED BY Sabeen**********************
            
         End If
         If FormName = "PrintPaymentOrder" Then
            rptFileName = App.Path & "\Reports\rptPaymentOrder.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableToolbar = True
            cmdPrint.Visible = True
            cmdVerify.Visible = False
         End If
         
         If FormName = "PaymentVoucher" Then
            rptFileName = App.Path & "\Reports\rptVoucher.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableToolbar = True
            cmdPrint.Visible = True
            cmdVerify.Visible = False
         End If
         
         '************MODIFIED BY Sabeen**********************
         If FormName = "frmSearchReceipts" Then
            rptFileName = App.Path & "\Reports\rptSearchReceipt.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableToolbar = True
            crvReport.EnableExportButton = True
            cmdPrint.Visible = True
            cmdVerify.Visible = False
         End If
         '****************************************************
         
         
         
        If FormName = "frmAFSClosingSourceOfFund" Then
            rptFileName = App.Path & "\Reports\rptSourceOfFundFinalization.rpt"
            crvReport.DisplayToolbar = True
            crvReport.EnableToolbar = True
            crvReport.EnableExportButton = True
            cmdPrint.Visible = True
            cmdVerify.Visible = False
            cmdVerify.Enabled = False
         End If
         
         
         arrInput = ArrayIn
         Screen.MousePointer = vbHourglass
         crvReport.DisplayTabs = True
         
         Set Rpt = Nothing
         mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
         Set Rpt = mApp.OpenReport(rptFileName, 1)
         
         If IsArray(arrInput) Then
             For mLoop = LBound(arrInput) To UBound(arrInput)
                 Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                 Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
             Next mLoop
         End If
         
         Screen.MousePointer = vbDefault
         crvReport.ReportSource = Rpt
         crvReport.ViewReport
         crvReport.Zoom (1)
    End Sub

    Public Property Let ArrayIn(mData As Variant)
        aryIn = mData
    End Property
    
    Public Property Get ArrayIn() As Variant
        ArrayIn = aryIn
    End Property
    
    Public Property Let ArrayString(mData As String)
        aryStr = mData
    End Property
    
    Public Property Get ArrayString() As String
        ArrayString = aryStr
    End Property
    Public Property Let FormName(mData As String)
        strFormName = mData
    End Property

    Public Property Get FormName() As String
        FormName = strFormName
    End Property
    
    Public Property Let MultipleVouchers(mData As Boolean)
        blnMultipleVouchers = mData
    End Property

    Public Property Get MultipleVouchers() As Boolean
        MultipleVouchers = blnMultipleVouchers
    End Property

