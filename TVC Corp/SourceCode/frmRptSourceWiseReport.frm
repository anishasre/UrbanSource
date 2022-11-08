VERSION 5.00
Begin VB.Form frmRptSourceWiseReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source of Fund wise Reports"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
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
   ScaleHeight     =   2910
   ScaleWidth      =   5340
   Begin VB.ComboBox cmbReportMenu 
      Height          =   360
      ItemData        =   "frmRptSourceWiseReport.frx":0000
      Left            =   90
      List            =   "frmRptSourceWiseReport.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   135
      Width           =   5100
   End
   Begin VB.Frame fmeMajorMinorHeadwise 
      Height          =   1950
      Left            =   90
      TabIndex        =   21
      Top             =   675
      Width           =   5190
      Begin VB.CommandButton cmdMajorMinorShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   1395
         Width           =   1095
      End
      Begin VB.TextBox txtMajorMinorTo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3465
         TabIndex        =   23
         Top             =   855
         Width           =   1500
      End
      Begin VB.TextBox txtMajorMinorFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         TabIndex        =   22
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "To Date:"
         Height          =   240
         Left            =   2790
         TabIndex        =   27
         Top             =   855
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
         Height          =   240
         Left            =   225
         TabIndex        =   26
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Major / Minor Account Head wise Receipt /Expenditure Statement From Own Fund"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   510
         Left            =   135
         TabIndex        =   25
         Top             =   180
         Width           =   4965
      End
   End
   Begin VB.Frame fmeSourcewiseCapitalAndRevenueExpendiutres 
      Height          =   1950
      Left            =   90
      TabIndex        =   14
      Top             =   675
      Width           =   5190
      Begin VB.TextBox txtCapitalFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         TabIndex        =   17
         Top             =   720
         Width           =   1500
      End
      Begin VB.TextBox txtCapitalTo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3465
         TabIndex        =   16
         Top             =   765
         Width           =   1500
      End
      Begin VB.CommandButton cmdCapitalShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2205
         TabIndex        =   15
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Statement Showing Capital and Revenue Expenditure from different"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   180
         Width           =   4965
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
         Height          =   240
         Left            =   225
         TabIndex        =   19
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "To Date:"
         Height          =   240
         Left            =   2790
         TabIndex        =   18
         Top             =   765
         Width           =   645
      End
   End
   Begin VB.Frame fmeSectorwise 
      Height          =   1950
      Left            =   90
      TabIndex        =   7
      Top             =   675
      Width           =   5190
      Begin VB.TextBox txtSectorFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
      Begin VB.TextBox txtSectorTo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3465
         TabIndex        =   9
         Top             =   765
         Width           =   1500
      End
      Begin VB.CommandButton cmdSectorShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2205
         TabIndex        =   8
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Sector wise Statement of Expenditure Devolopment Fund"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   135
         TabIndex        =   13
         Top             =   180
         Width           =   4965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
         Height          =   240
         Left            =   225
         TabIndex        =   12
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To Date:"
         Height          =   240
         Left            =   2790
         TabIndex        =   11
         Top             =   765
         Width           =   645
      End
   End
   Begin VB.Frame fmeSourcewiseReceiptAndPaymentStatement 
      Height          =   1950
      Left            =   90
      TabIndex        =   0
      Top             =   675
      Width           =   5190
      Begin VB.CommandButton cmdRPShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2205
         TabIndex        =   6
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtRPTo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3465
         TabIndex        =   5
         Top             =   765
         Width           =   1500
      End
      Begin VB.TextBox txtRPFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         TabIndex        =   3
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To Date:"
         Height          =   240
         Left            =   2790
         TabIndex        =   4
         Top             =   765
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sorce of Fund wise Receipt and Payment Statement"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   4965
      End
   End
End
Attribute VB_Name = "frmRptSourceWiseReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mAryInput As Variant

    Private Sub cmbReportMenu_Click()
        Call VisibleFalse
        If cmbReportMenu.ListIndex > 2 Then
            If cmbReportMenu.ListIndex = 3 Then
                mAryInput = Array(10, 1)
            ElseIf cmbReportMenu.ListIndex = 4 Then
                mAryInput = Array(10, 2)
            ElseIf cmbReportMenu.ListIndex = 5 Then
                mAryInput = Array(20, 1)
            ElseIf cmbReportMenu.ListIndex = 6 Then
                mAryInput = Array(20, 2)
            End If
            fmeMajorMinorHeadwise.Visible = True
        Else
            If cmbReportMenu.ListIndex = 0 Then
                fmeSourcewiseReceiptAndPaymentStatement.Visible = True
            ElseIf cmbReportMenu.ListIndex = 1 Then
                fmeSectorwise.Visible = True
            ElseIf cmbReportMenu.ListIndex = 2 Then
                fmeSourcewiseCapitalAndRevenueExpendiutres.Visible = True
            End If
        End If
    End Sub

    Private Sub cmdCapitalShow_Click()
        mAryInput = Array(CDate(txtCapitalFrom.Text), CDate(txtCapitalTo.Text))
        ShowReport mAryInput, "rptCapitalAndRevenueStatement.rpt"
    End Sub

    Private Sub cmdMajorMinorShow_Click()
'        mAryInput(2) = CDate(txtMajorMinorFrom.Text)
'        mAryInput(3) = CDate(txtMajorMinorTo.Text)
        mAryInput = Array(mAryInput(0), mAryInput(1), CDate(txtMajorMinorFrom.Text), CDate(txtMajorMinorTo.Text))
        ShowReport mAryInput, "rptSourcewiseMajorMinorHeads.rpt"
    End Sub

    Private Sub cmdRPShow_Click()
        mAryInput = Array(CDate(txtRPFrom.Text), CDate(txtRPTo.Text))
        ShowReport mAryInput, "rptSourcewiseRPStatement.rpt"
    End Sub

    Private Sub cmdSectorShow_Click()
        mAryInput = Array(CDate(txtSectorFrom.Text), CDate(txtSectorTo.Text))
        ShowReport mAryInput, "rptSectorwiseExpenditure.rpt"
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub

    Private Sub Form_Load()
        Call FormInitialise
    End Sub
    Private Sub FormInitialise()
        txtRPFrom.Text = DdMmmYy(gbStartingDate)
        txtSectorFrom.Text = DdMmmYy(gbStartingDate)
        txtCapitalFrom.Text = DdMmmYy(gbStartingDate)
        txtMajorMinorFrom.Text = DdMmmYy(gbStartingDate)
        
        txtRPTo.Text = DdMmmYy(gbTransactionDate)
        txtCapitalTo.Text = DdMmmYy(gbTransactionDate)
        txtMajorMinorTo.Text = DdMmmYy(gbTransactionDate)
        txtSectorTo.Text = DdMmmYy(gbTransactionDate)
    End Sub

    Private Sub txtRPFrom_LostFocus()
        txtRPFrom.Text = DdMmmYy(Trim(txtRPFrom.Text))
    End Sub
    
    Private Sub txtRPTo_LostFocus()
        txtRPTo.Text = DdMmmYy(Trim(txtRPTo.Text))
    End Sub
    
    Private Sub txtSectorFrom_LostFocus()
        txtSectorFrom.Text = DdMmmYy(Trim(txtSectorFrom.Text))
    End Sub
    
    Private Sub txtSectorTo_LostFocus()
        txtSectorTo.Text = DdMmmYy(Trim(txtSectorTo.Text))
    End Sub
    
    Private Sub txtCapitalFrom_LostFocus()
        txtCapitalFrom.Text = DdMmmYy(Trim(txtCapitalFrom.Text))
    End Sub
    
    Private Sub txtCapitalTo_LostFocus()
        txtCapitalTo.Text = DdMmmYy(Trim(txtCapitalTo.Text))
    End Sub
    
    Private Sub txtMajorMinorFrom_LostFocus()
        txtMajorMinorFrom.Text = DdMmmYy(Trim(txtMajorMinorFrom.Text))
    End Sub
    
    Private Sub txtMajorMinorTo_LostFocus()
        txtMajorMinorTo.Text = DdMmmYy(Trim(txtMajorMinorTo.Text))
    End Sub
    
    Private Sub ShowReport(aryIn As Variant, ReportName As String)
        Dim mLoop As Integer
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        
        arInput = aryIn
        frmNewRpt.rptFileName = App.Path & "\Reports\" + ReportName
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub
    
    Private Sub VisibleFalse()
        fmeMajorMinorHeadwise.Visible = False
        fmeSectorwise.Visible = False
        fmeSourcewiseCapitalAndRevenueExpendiutres.Visible = False
        fmeSourcewiseReceiptAndPaymentStatement.Visible = False
    End Sub
