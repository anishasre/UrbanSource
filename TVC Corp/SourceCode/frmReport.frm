VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8325
      lastProp        =   500
      _cx             =   14684
      _cy             =   10663
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    CRV.Left = 0
    CRV.Top = 0
    CRV.Width = Me.Width
    CRV.Height = Me.Height
    showreport "rptJournal1.rpt"
End Sub

Private Function showreport(rptName)
    Dim aryInputRpt(1) As Variant
    Dim Aryout As Variant
    Dim rpt1 As New CRAXDRT.Report
    Dim app1 As New CRAXDRT.Application
    aryInputRpt(0) = 1
    aryInputRpt(1) = 40
    Screen.MousePointer = vbHourglass
    CRV.EnableExportButton = True
    Set rpt1 = Nothing
    app1.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
    Set rpt1 = app1.OpenReport(App.Path & "\Reports\" & rptName, 1)
    rpt1.ParameterFields.Item(1).AddCurrentValue (aryInputRpt(0))
    rpt1.ParameterFields.Item(2).AddCurrentValue (aryInputRpt(1))
    CRV.EnableProgressControl = True
    CRV.EnableProgressControl = True
    CRV.DisplayGroupTree = False
    Screen.MousePointer = vbDefault
    CRV.ReportSource = rpt1
    CRV.Refresh
    CRV.ViewReport
End Function

Private Sub Form_Resize()
    CRV.Left = 0
    CRV.Top = 0
    CRV.Width = Me.Width
    CRV.Height = Me.Height
End Sub
