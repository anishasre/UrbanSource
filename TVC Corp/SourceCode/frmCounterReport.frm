VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCounterReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saankhya - Counter Reports"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Report"
      Height          =   315
      Left            =   5730
      TabIndex        =   14
      Top             =   2400
      Width           =   1185
   End
   Begin WinXPC_Engine.WindowsXPC WinXPC 
      Left            =   -3510
      Top             =   2310
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fmeOperator 
      Height          =   2265
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cmbSection 
         Enabled         =   0   'False
         Height          =   360
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   2415
      End
      Begin VB.ComboBox cmbUser 
         Height          =   360
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbSession 
         Height          =   360
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cmbShift 
         Height          =   360
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1140
         Width           =   2415
      End
      Begin VB.ComboBox cmbCounter 
         Height          =   360
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   345
         Left            =   810
         TabIndex        =   1
         Top             =   750
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         Format          =   59310081
         CurrentDate     =   39719
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section :"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   270
         Left            =   5460
         TabIndex        =   12
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Session :"
         Height          =   240
         Left            =   3630
         TabIndex        =   11
         Top             =   1590
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift :"
         Height          =   240
         Left            =   3840
         TabIndex        =   10
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         Height          =   240
         Left            =   3390
         TabIndex        =   9
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter :"
         Height          =   240
         Left            =   60
         TabIndex        =   8
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   240
         Left            =   330
         TabIndex        =   7
         Top             =   840
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCounterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mCounter As Variant
    Dim mUserID As Variant
    Dim mShift As Variant
    Dim mSession As Variant
    Private Sub cmbCounter_Click()
        If cmbCounter.ListIndex > 0 Then
            PopulateList cmbUser, "Select Distinct A.vchUserName,A.numUserID From faVouchers Inner Join faUser A On A.numUserID=faVouchers.intUserID Where tnyVoucherTypeID=10 and dtDate= '" & CheckDateInMMM(dtpDate.value) & "' And intCounterID = " & cmbCounter.ItemData(cmbCounter.ListIndex), , True, , True
            mCounter = CStr(cmbCounter.ItemData(cmbCounter.ListIndex))
        Else
            cmbUser.Clear
            mCounter = "%"
        End If
    End Sub

    Private Sub cmbSection_Click()
        If cmbSection.ListIndex > 0 Then
            PopulateList cmbCounter, "Select Distinct faCounters.vchDesCription,faCounters.intCounterID From faVouchers Inner Join faCounters On faCounters.intCounterID=faVouchers.intCounterID Where tnyVoucherTypeID=10 And dtDate= '" & CheckDateInMMM(dtpDate.value) & "'", , True, , True
        Else
            cmbCounter.Clear
        End If
    End Sub
    
    Private Sub cmbSession_Click()
        If cmbSession.ListIndex > 0 Then
            mSession = CStr(cmbSession.ItemData(cmbSession.ListIndex))
        Else
            mSession = "%"
        End If
    End Sub

    Private Sub cmbShift_Click()
        If cmbShift.ListIndex > 0 Then
            PopulateList cmbSession, "Select Distinct 'Session '+CONVERT(varchar(10),intSessionID),intSessionID From faVouchers Where tnyVoucherTypeID=10 and dtDate= '" & CheckDateInMMM(dtpDate.value) & "' And intCounterID = " & cmbCounter.ItemData(cmbCounter.ListIndex) & " And intUserID = " & cmbUser.ItemData(cmbUser.ListIndex) & " And tnyShiftID = " & cmbShift.ItemData(cmbShift.ListIndex), , True, , True
            mShift = CStr(cmbShift.ItemData(cmbShift.ListIndex))
        Else
            cmbSession.Clear
            mShift = "%"
        End If
    End Sub

    Private Sub cmbUser_Click()
        If cmbUser.ListIndex > 0 Then
            PopulateList cmbShift, "Select Distinct faShifts.vchShift,faShifts.intShiftID From faVouchers Inner Join faShifts On faShifts.intShiftID=faVouchers.tnyShiftID Where tnyVoucherTypeID=10 and dtDate= '" & CheckDateInMMM(dtpDate.value) & "' And intCounterID = " & cmbCounter.ItemData(cmbCounter.ListIndex) & " And intUserID = " & cmbUser.ItemData(cmbUser.ListIndex), , True, , True
            mUserID = CStr(cmbUser.ItemData(cmbUser.ListIndex))
        Else
            cmbShift.Clear
            mUserID = "%"
        End If
    End Sub

    Private Sub cmdShow_Click()
        Dim objDb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        ''frmMenu.Transactions.Enabled = False
        arInput = Array(dtpDate.value, mCounter, mUserID, mShift, mSession)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCounterReports.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
        gbCounterStatusFlag = True
    End Sub

    Private Sub Form_Load()
        WinXPC.InitIDESubClassing
        lblDate.Caption = CheckDateInMMM(Date)
        dtpDate.value = Date
        '---------------------------------------------------------------------'
        '                           Populate List                             '
        '---------------------------------------------------------------------'
        PopulateList cmbSection, "Select vchSectionName,intSectionID From faSection", , True, True, True
        If gbLBType = 3 Or gbLBType = 4 Then
            cmbSection.Text = "Janasevana Kendram"
        Else
            cmbSection.Text = "Front Office"
        End If
        '---------------------------------------------------------------------'
        mUserID = "%"
        mCounter = "%"
        mShift = "%"
        mSession = "%"
        If gbUserTypeID = 4 Then
            On Error Resume Next
            fmeOperator.Enabled = True
            cmbCounter.Text = gbCounterName
            cmbUser.Text = gbUserName
            cmbShift.Text = gbShiftName ' Newly Added
            cmbSession.Text = "Session " & val(gbSessionID)
            On Error GoTo 0
        ElseIf gbUserTypeID = 1 Then
            fmeOperator.Enabled = True
        End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        frmCounterReport.Visible = False
    End Sub
