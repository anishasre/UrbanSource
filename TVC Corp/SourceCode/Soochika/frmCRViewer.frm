VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmCRViewer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7155
   ClientLeft      =   750
   ClientTop       =   2535
   ClientWidth     =   9960
   Icon            =   "frmCRViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnPrint 
      BackColor       =   &H0080C0FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8490
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9810
      UseMaskColor    =   -1  'True
      Width           =   885
   End
   Begin VB.CommandButton btnClose 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9420
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9810
      UseMaskColor    =   -1  'True
      Width           =   885
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CR 
      Height          =   8055
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      lastProp        =   500
      _cx             =   17595
      _cy             =   14208
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
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
Attribute VB_Name = "frmCRViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objApp As New CRAXDRT.Application
 
Public objrpt As CRAXDRT.Report



Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub btnPrint_Click()
    CR.PrintReport
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    CR.Left = 0
    CR.Top = 0
    CR.Width = Me.Width - 175
    CR.Height = Me.Height - 450
End Sub
Public Function vShowReport(ByVal sReportPath As String, ByVal sReportName As String, Optional vParamArray As Variant)
    On Error GoTo ErrorHandler
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    Set objrpt = objApp.OpenReport(sReportPath & "\" & sReportName, 1)
    objApp.LogOnServer "ODBC", "GWFlow1", "DB_SoochikaUrban", "dbsoochika", "urbansoochika"
    'objrpt.PrinterName=
    If Not IsMissing(vParamArray) Then
        'objRpt.ParameterFields.Item(1).ClearCurrentValueAndRange
        For i = 0 To UBound(vParamArray) - 1
            objrpt.ParameterFields.Item(i + 1).ClearCurrentValueAndRange
            objrpt.ParameterFields(i + 1).AddCurrentValue vParamArray(i)
        Next
    End If
    
    'Added for testing Starts
    'objrpt.PageEngine.
    'myreport.PrintOptions.PaperSize = New CrystalDecisions.Shared.PaperSize
    'Ends
    
    CR.ReportSource = objrpt
    CR.ViewReport
    CR.Zoom (150)
    Screen.MousePointer = vbDefault
    Set objrpt = Nothing
    Set objApp = Nothing
    Exit Function
ErrorHandler:
    MsgBox VBA.Err.Number & ":  " & VBA.Err.Description, vbCritical + vbOKOnly
    Resume Next
End Function
Public Function vShowReportOLEDB(ByVal sReportPath As String, ByVal sReportName As String, ByRef objConnection As ADODB.Connection, Optional vParamArray As Variant)
   ' On Error GoTo ErrorHandler
    Dim i As Integer
    Dim strServer As String, strDataBase As String, strUserName As String, strPassWord As String
    strServer = objConnection.Properties("Server Name")
    strDataBase = objConnection.Properties("Current Catalog")
    strUserName = objConnection.Properties("User ID")
    strPassWord = objConnection.Properties("Password")
    Dim CPProperties As CRAXDRT.ConnectionProperties
    Dim DBTable As CRAXDRT.DatabaseTable
    Screen.MousePointer = vbHourglass
    
    Set objrpt = objApp.OpenReport(sReportPath & "\" & sReportName, 1)
    Set DBTable = objrpt.Database.Tables(1)
    Set CPProperties = DBTable.ConnectionProperties
    CPProperties.DeleteAll
    If Not IsMissing(vParamArray) Then
        For i = 0 To UBound(vParamArray) - 1
            objrpt.ParameterFields.Item(i + 1).ClearCurrentValueAndRange
            objrpt.ParameterFields(i + 1).AddCurrentValue vParamArray(i)
        Next
    End If
    
    CPProperties.Add "Provider", "SQLOLEDB"
    CPProperties.Add "Data Source", strServer
    CPProperties.Add "Initial Catalog", strDataBase
    CPProperties.Add "User ID", strUserName
    CPProperties.Add "Password", Trim(strPassWord)
    CR.ReportSource = objrpt
    CR.ViewReport
    Screen.MousePointer = vbDefault
    Set objrpt = Nothing
    Set objApp = Nothing
    Exit Function
ErrorHandler:
    MsgBox VBA.Err.Number & ":  " & VBA.Err.Description, vbCritical + vbOKOnly
    Resume Next
End Function

Public Function ShowUnicodeReport(ByVal sReportPath As String, ByVal sReportName As String, Optional vParamArray As Variant)
    On Error GoTo ErrorHandler
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    Set objrpt = objApp.OpenReport(sReportPath & "\" & sReportName, 1)
    objApp.LogOnServer "ODBC", "GWFlow1", "DB_Soochika", "workflow", "A+v378*R"
    'objrpt.PrinterName=
    If Not IsMissing(vParamArray) Then
        'objRpt.ParameterFields.Item(1).ClearCurrentValueAndRange
        For i = 0 To UBound(vParamArray) - 1
            objrpt.ParameterFields.Item(i + 1).ClearCurrentValueAndRange
            objrpt.ParameterFields(i + 1).AddCurrentValue vParamArray(i)
        Next
    End If
    
    'Added for testing Starts
    'objrpt.PageEngine.
    'myreport.PrintOptions.PaperSize = New CrystalDecisions.Shared.PaperSize
    'Ends
    
    CR.ReportSource = objrpt
    CR.ViewReport
     CR.EnableExportButton = True
    CR.Zoom (150)
    Screen.MousePointer = vbDefault
    Set objrpt = Nothing
    Set objApp = Nothing
    Exit Function
ErrorHandler:
    MsgBox VBA.Err.Number & ":  " & VBA.Err.Description, vbCritical + vbOKOnly
    Resume Next
End Function
