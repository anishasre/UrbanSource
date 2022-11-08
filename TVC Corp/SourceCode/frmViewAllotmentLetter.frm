VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewAllotmentLetter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Allotment Letter"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5070
      TabIndex        =   7
      Top             =   8760
      Width           =   1890
   End
   Begin VB.CommandButton cmdViewLetterOfAllotment 
      Caption         =   "View &Letter of Allotment"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   6
      Top             =   8775
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14250
      TabIndex        =   1
      Top             =   9180
      Width           =   14310
      Begin VB.CommandButton cmdChangeProceedings 
         Caption         =   "Change Proceedings No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3390
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.CommandButton cmdViewAuthorizationReport 
         Caption         =   "View &Authorization Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   5
         Top             =   90
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.CommandButton cmdViewRequisitionReport 
         Caption         =   "View &Requisition Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   4
         Top             =   90
         Visible         =   0   'False
         Width           =   2370
      End
      Begin VB.CheckBox cmdVerify 
         Caption         =   "&Verify Requisition"
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
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1890
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
         TabIndex        =   2
         Top             =   60
         Width           =   1725
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   8700
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   14295
      lastProp        =   500
      _cx             =   25215
      _cy             =   15346
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmViewAllotmentLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim aryIn   As Variant
    Dim mMode   As Variant
    
    Dim mLoadModeUnAuth       As Integer      '10-For UNAUTHORIZED DRAWAL
    Dim mPreviousYearsRequestID As Integer
    
    '*********************************************************************************************'
    '                           Form to view the Reports related to Requisition                   '
    '*********************************************************************************************'
    Private Sub ReportView()
         Dim Rpt As New CRAXDRT.Report
         Dim mApp As New CRAXDRT.Application
         Dim rptFileName As String
         Dim arrInput As Variant
         Dim mLoop As Long
         
         cmdChangeProceedings.Visible = False
         If Mode = 1 Then
'            If gbUserTypeID <> 3 Then
'                cmdVerify.Visible = True
'            Else
'                cmdVerify.Visible = False
'            End If
            frmViewAllotmentLetter.Caption = "View Allotment Letter"
            rptFileName = App.Path & "\Reports\rptViewAllotmentLetter.rpt"
         ElseIf Mode = 2 Then
            frmViewAllotmentLetter.Caption = "View Authorization Letter"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = True
            rptFileName = App.Path & "\Reports\rptViewAuthorizationLetter.rpt"
        ElseIf Mode = 3 Then
            frmViewAllotmentLetter.Caption = "Letter of Allotment"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptLetterOfAllotment.rpt"
        ElseIf Mode = 4 Then
            cmdChangeProceedings.Visible = True
            frmViewAllotmentLetter.Caption = "Proceedings"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptProceedingsGEN-40.rpt"
        ElseIf Mode = 5 Then
            frmViewAllotmentLetter.Caption = "Treasury Bill (TR 59 A)"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptTreasuryBill-59(C)-BFund.rpt" 'rptTreasuryBillGEN-43.rpt"
                                              
        ElseIf Mode = 6 Then '::TR59(B)_59(C)  CHANGED TO TR 59 B
            frmViewAllotmentLetter.Caption = "Treasury Bill (TR 59 B)"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptTreasuryBill-59(C)_59(B).rpt" 'rptTreasuryBillGEN-44.rpt"
        ElseIf Mode = 7 Then
            frmViewAllotmentLetter.Caption = "Appropriation Contriol Register"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptAppropriationControlRegisterGEN-39.rpt"
        ElseIf Mode = 8 Then
            frmViewAllotmentLetter.Caption = "Treasury Bill (TR 59 A)"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptTreasuryBill-59(C) TSB.rpt" 'rptTreasuryBill_TSB.rpt"
        ElseIf Mode = 9 Then
            ' Treasury Bill - New Mode - TR59(c)
            frmViewAllotmentLetter.Caption = "Treasury Bill TR-59(C)"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptTreasuryBill-59(C).rpt" 'rptTreasuryBillGEN-44 CON FUND.rpt
        ElseIf Mode = 10 Then
            ' "TRANSFER CREDIT"
            frmViewAllotmentLetter.Caption = "TRANSFER CREDIT"
            cmdVerify.Visible = False
            cmdViewRequisitionReport.Visible = False
            cmdViewAuthorizationReport.Visible = False
            rptFileName = App.Path & "\Reports\rptTransferCredit.rpt"
        End If
         arrInput = ArrayIn
         Screen.MousePointer = vbHourglass
         crvReport.DisplayToolbar = True
         
         Set Rpt = Nothing
         mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
         Set Rpt = mApp.OpenReport(rptFileName, 1)
         
         If IsArray(arrInput) Then
             For mLoop = LBound(arrInput) To UBound(arrInput)
                 Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                 If Mode <> 3 And Mode <> 1 Then
                    Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
                 Else
                    Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue Trim(str(arrInput(mLoop)))
                 End If
             Next mLoop
         End If
         Screen.MousePointer = vbDefault
         crvReport.ReportSource = Rpt
         crvReport.ViewReport
         crvReport.Zoom (1)
    End Sub
    
    Private Sub ProceedingNumber()
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        
        gbSearchID = -1
        gbSearchStr = ""
        frmProceedings.chkEdit.value = 0
        frmProceedings.Module = 130
        frmProceedings.Show vbModal
        If gbSearchID > 0 Then
            Dim objProceedings As New clsProceedings
            With objProceedings
                .ProceedingsID = gbSearchID
                .getProceedingsByID
                If .Used > 0 Then
                    MsgBox "This Proceedings already used", vbInformation
                    .ProceedingsID = -1
                Else
                    
                    mSQL = " DELETE FROM faProceedings WHERE intModuleID = 130 AND intVoucherID = " & val(ArrayIn(0))
                    objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                    
                    mSQL = "UPDATE  faProceedings SET intModuleID = 130, tnyUsed=1 , intVoucherID = " & val(ArrayIn(0))
                    mSQL = mSQL + " Where intProceedingsID= " & gbSearchID
                    objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                End If
            End With
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub
        
    Private Sub cmdChangeProceedings_Click()
        Call ProceedingNumber
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
'        frmListOfRequisitions.Visible = True
'        frmListOfRequisitions.ZOrder (0)
    End Sub

    Private Sub cmdPrint_Click()
        crvReport.PrintReport
    End Sub
 Public Function CheckPreviousYearRequisitions(mReqID As Variant) As Integer
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mReqID <> "" Then
            If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSQL = "Select * From faPendingTaskRequest Where intKeyID=" & mReqID
                Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    CheckPreviousYearRequisitions = 1
                    mPreviousYearsRequestID = Rec!intRequestID
                Else
                    CheckPreviousYearRequisitions = 0
                    mPreviousYearsRequestID = -1
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Function

    Private Sub cmdVerify_Click()
        'Unload Me
        Dim mPreviousYearMode As Integer
        frmRequisition.RequisitionID = ArrayIn(0)
        
        ' NOTE:: To check and set previous year mode
        '    :: Modified by Aiby (14-Jul-2013)
        If ArrayIn(1) < gbFinancialYearID Then
            mPreviousYearMode = CheckPreviousYearRequisitions(ArrayIn(0))
            frmRequisition.PreviousYearMode = mPreviousYearMode
            frmRequisition.PreviousYearRequestID = mPreviousYearsRequestID ' ##28-May-2014 [AIBY] [MODIFIED FOR COCHINCORP]
            If mLoadModeUnAuth = 10 Then
                frmRequisition.LoadMode = 10
            End If
        End If
        Unload Me
        frmRequisition.Show vbModal
        
    End Sub

    Private Sub cmdViewAuthorizationReport_Click()
        Mode = 2
        Call ReportView
        cmdViewRequisitionReport.Visible = True
        cmdViewAuthorizationReport.Visible = False
    End Sub

    Private Sub cmdViewLetterOfAllotment_Click()
        Mode = 3
        Call ReportView
    End Sub

    Private Sub cmdViewRequisitionReport_Click()
        Mode = 1
        Call ReportView
        cmdViewAuthorizationReport.Visible = True
        cmdViewRequisitionReport.Visible = False
    End Sub

    Private Sub Form_Load()
        Call ReportView
    End Sub
    'Property for getting the array input for the Report
    Public Property Let ArrayIn(mData As Variant)
        aryIn = mData
    End Property
    
    Public Property Get ArrayIn() As Variant
        ArrayIn = aryIn
    End Property
    'Property for getting the input for which Report to show
    Public Property Let Mode(mData As Variant)
        mMode = mData
    End Property
    
    Public Property Get Mode() As Variant
        Mode = mMode
    End Property
    Private Sub Form_Unload(Cancel As Integer)
    If mMode = 8 Then
        frmListOfAllotments.Visible = True
        frmListOfAllotments.ZOrder (0)
    ElseIf mMode = 10 Then
        frmContraEntry.Visible = True
        frmContraEntry.ZOrder (0)
    Else
        If mLoadModeUnAuth = 10 Then
            frmListOfRequisitions.LoadMode = 10
            frmListOfRequisitions.Visible = True
            frmListOfRequisitions.ZOrder (0)
            mLoadModeUnAuth = 0
        Else
            frmListOfRequisitions.Visible = True
            frmListOfRequisitions.ZOrder (0)
        End If
    End If
    End Sub
    
    Public Property Let LoadMode(mData As Integer)
        mLoadModeUnAuth = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadModeUnAuth
    End Property
