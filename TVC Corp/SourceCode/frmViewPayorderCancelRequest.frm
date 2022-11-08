VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewPayorderCancelRequest 
   BorderStyle     =   0  'None
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12195
      TabIndex        =   0
      Top             =   6840
      Width           =   12255
      Begin VB.CheckBox cmdVerify 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8610
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1725
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
         TabIndex        =   1
         Top             =   60
         Width           =   1725
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   6765
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   12240
      lastProp        =   500
      _cx             =   21590
      _cy             =   11933
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
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
Attribute VB_Name = "frmViewPayorderCancelRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private aryIn As Variant
    Private mCancelled As Boolean

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdVerify_Click()
        Unload Me
        frmViewPaymentorderCancellationRequest.Verified = True
        frmViewPaymentorderCancellationRequest.cmdView.Caption = "Cancel"
        frmViewPaymentorderCancellationRequest.Show vbModal
    End Sub

    Private Sub Form_Load()
        Call ReportView
    End Sub
    Private Sub ReportView()
        Dim Rpt As New CRAXDRT.Report
        Dim mApp As New CRAXDRT.Application
        Dim rptFileName As String
        Dim arrInput As Variant
        Dim mLoop As Long
         
        rptFileName = App.Path & "\Reports\rptPayOrderCancelRequests.rpt"
        crvReport.DisplayToolbar = True
        crvReport.EnableNavigationControls = True
        crvReport.EnableToolbar = True

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
   
    Public Property Let Cancellation(mData As Boolean)
        mCancelled = mData
    End Property
    
    Public Property Let PayOrderNo(mData As Variant)
    
    End Property

