VERSION 5.00
Begin VB.Form frmPropertyTaxCalculator 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Property Tax Calculator"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNonR 
      BackColor       =   &H00DAF2F2&
      Caption         =   "NonResidential"
      Height          =   240
      Left            =   315
      TabIndex        =   26
      Top             =   90
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkSTax 
      Caption         =   "ServiceCess"
      Height          =   195
      Left            =   4275
      TabIndex        =   25
      Top             =   4140
      Width           =   1185
   End
   Begin VB.CheckBox chkFineWaiver 
      Caption         =   "FineWaiver"
      Height          =   225
      Left            =   4200
      TabIndex        =   24
      Top             =   2745
      Width           =   1170
   End
   Begin VB.CommandButton cmdCopyToReceipt 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Copy to Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4035
      Width           =   1620
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2055
      Width           =   1440
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DAF2F2&
      Height          =   795
      Left            =   270
      TabIndex        =   10
      Top             =   450
      Width           =   6195
      Begin VB.OptionButton optFullYearRate 
         BackColor       =   &H00DAF2F2&
         Caption         =   "&Full Year"
         Enabled         =   0   'False
         Height          =   210
         Left            =   3585
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   1125
      End
      Begin VB.OptionButton optHalfYearRate 
         BackColor       =   &H00DAF2F2&
         Caption         =   "&Half Year"
         Height          =   210
         Left            =   3585
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.TextBox txtTaxRate 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1575
         TabIndex        =   1
         Top             =   270
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Half Year Tax:"
         Height          =   195
         Left            =   105
         TabIndex        =   0
         Top             =   300
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DAF2F2&
      Height          =   630
      Left            =   3360
      TabIndex        =   12
      Top             =   1275
      Width           =   3120
      Begin VB.TextBox txtToPeriodID 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2010
         MaxLength       =   1
         TabIndex        =   8
         Top             =   195
         Width           =   480
      End
      Begin VB.TextBox txtToYear 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   510
         MaxLength       =   4
         TabIndex        =   7
         Top             =   195
         Width           =   1320
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1860
         TabIndex        =   20
         Top             =   90
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAF2F2&
      Height          =   615
      Left            =   270
      TabIndex        =   11
      Top             =   1290
      Width           =   3045
      Begin VB.TextBox txtFromPeriodID 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   6
         Top             =   180
         Width           =   480
      End
      Begin VB.TextBox txtFromYear 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   510
         MaxLength       =   4
         TabIndex        =   5
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1830
         TabIndex        =   19
         Top             =   90
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   345
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fine"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   3330
      Width           =   705
   End
   Begin VB.Label lblFine 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   21
      Top             =   3315
      Width           =   1440
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   255
      TabIndex        =   18
      Top             =   3675
      Width           =   6195
   End
   Begin VB.Label lblLC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   17
      Top             =   3015
      Width           =   1440
   End
   Begin VB.Label lblPT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   16
      Top             =   2730
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Cess"
      Height          =   195
      Left            =   1695
      TabIndex        =   15
      Top             =   3030
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property Tax"
      Height          =   195
      Left            =   1665
      TabIndex        =   14
      Top             =   2745
      Width           =   900
   End
End
Attribute VB_Name = "frmPropertyTaxCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mNoOfHalfYears As Integer
    Dim mDemandMode As Boolean
    Private mMode As Integer
    Private mPTRow As Integer
    Private mNonR As Integer
    Private mHYrTaxInFraction As Boolean
    
    Private Sub copyToDemand()
        Dim mLoop As Integer
        Dim mYearID As Integer
        Dim mPeriodID As Integer
        Dim objAcc As New clsAccounts
        Dim mRow As Integer
        Dim mFineFlag As Boolean
        
        mYearID = val(txtFromYear)
        mPeriodID = val(txtFromPeriodID)
        mRow = 1
        frmDemandInterface.vsGrid.Clear 1, 0
        For mLoop = 1 To mNoOfHalfYears
            If mRow > 8 Then
                'MsgBox "Can't Print more than 8 Rows...", vbInformation
                'Exit For
            End If
            If mYearID < gbFinancialYearID Then
                mFineFlag = True
                objAcc.SetAccountCode gbAcHeadCodePropertyTaxArrear
            Else
                objAcc.SetAccountCode gbAcHeadCodePropertyTaxCurrent
            End If
            If frmDemandInterface.vsGrid.Rows = mRow Then
                frmDemandInterface.vsGrid.Rows = frmDemandInterface.vsGrid.Rows + 10
            End If
            
            If objAcc.AccountHeadID > 0 Then
                frmDemandInterface.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                frmDemandInterface.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                frmDemandInterface.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                
                If mPeriodID = 1 Then
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 3) = 1 '"First Half"
                Else
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 3) = 2 '"Second Half"
                End If
                If mYearID < gbFinancialYearID Then
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 4) = val(lblPT.Caption)
                Else
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 5) = val(lblPT.Caption)
                End If
                frmDemandInterface.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 7) = mYearID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 9) = 1
                frmDemandInterface.vsGrid.TextMatrix(mRow, 10) = ""
                frmDemandInterface.vsGrid.TextMatrix(mRow, 11) = val(lblPT.Caption)
                frmDemandInterface.vsGrid.TextMatrix(mRow, 12) = ""
            End If
            
            mRow = mRow + 1
            If frmDemandInterface.vsGrid.Rows = mRow Then
                frmDemandInterface.vsGrid.Rows = frmDemandInterface.vsGrid.Rows + 2
            End If
            
            objAcc.SetAccountCode gbAcHeadCodeLibraryCess
            If objAcc.AccountHeadID > 0 Then
                frmDemandInterface.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                frmDemandInterface.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                frmDemandInterface.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                If mPeriodID = 1 Then
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 3) = "First Half"
                Else
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 3) = "Second Half"
                End If
                If mYearID < gbFinancialYearID Then
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 4) = val(lblLC.Caption)
                Else
                    frmDemandInterface.vsGrid.TextMatrix(mRow, 5) = val(lblLC.Caption)
                End If
                frmDemandInterface.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 7) = mYearID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 9) = 1
                frmDemandInterface.vsGrid.TextMatrix(mRow, 10) = ""
                frmDemandInterface.vsGrid.TextMatrix(mRow, 11) = val(lblLC.Caption)
                frmDemandInterface.vsGrid.TextMatrix(mRow, 12) = ""
            End If
            
            mRow = mRow + 1
            If frmDemandInterface.vsGrid.Rows = mRow Then
                frmDemandInterface.vsGrid.Rows = frmDemandInterface.vsGrid.Rows + 2
            End If
        
            If mPeriodID = 1 Then
                mPeriodID = 2
            Else
                mPeriodID = 1
                mYearID = mYearID + 1
            End If
        Next mLoop
        'If mFineFlag Then
         If val(lblFine.Caption) > 0 Then
            objAcc.SetAccountCode gbAcHeadCodePenalInterest
            If objAcc.AccountHeadID > 0 Then
                frmDemandInterface.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                frmDemandInterface.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                frmDemandInterface.vsGrid.TextMatrix(mRow, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
                frmDemandInterface.vsGrid.TextMatrix(mRow, 3) = ""
            
                frmDemandInterface.vsGrid.TextMatrix(mRow, 5) = val(lblFine.Caption)
                frmDemandInterface.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 7) = gbFinancialYearID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                frmDemandInterface.vsGrid.TextMatrix(mRow, 9) = 1
                frmDemandInterface.vsGrid.TextMatrix(mRow, 10) = ""
                frmDemandInterface.vsGrid.TextMatrix(mRow, 11) = val(lblFine.Caption)
                frmDemandInterface.vsGrid.TextMatrix(mRow, 12) = ""
            End If
        End If
        'End If
        frmDemandInterface.Calculate
        Unload Me
    End Sub
    Private Function CalculateFineforPTax(mYearID As Integer, mPeriodID As Integer, mPTax As Double) As Double
        '==============================================================================='
        ' Modified By : Aiby                                                            '
        '             : For Calicut Corporation                                         '
        '                                                                               '
        '==============================================================================='
        Dim dtFromDt As Variant
        Dim mNoOfMonths As Integer
        Dim mAmount     As Double
        Dim mFineAmt    As Double
        Dim mYearDiff   As Integer
        Dim mRate(1994 To 2009) As Integer
        Dim mFineRate As Integer
        
        Dim mTotalFineAmt As Double
        
        mRate(1994) = 288
        mRate(1995) = 264
        mRate(1996) = 240
        mRate(1997) = 216
        mRate(1998) = 208
        mRate(1999) = 208
        mRate(2000) = 208
        mRate(2001) = 184
        mRate(2002) = 160
        mRate(2003) = 136
        mRate(2004) = 112
        mRate(2005) = 88
        mRate(2006) = 64
        mRate(2007) = 40
        mRate(2008) = 24
        mRate(2009) = 12
      
        
        If mYearID <> 2010 Then
            If (mYearID) < 1994 Then
                mFineRate = mRate(1994) / 2
            Else
                mFineRate = mRate(mYearID) / 2
            End If
            mFineAmt = mPTax * mFineRate / 100
        Else
            mFineAmt = 0
        End If
        CalculateFineforPTax = mFineAmt
    End Function
    
    Private Function Fine(ByVal mYearID As Integer, ByVal mPeriodID As Integer, ByVal mUptoDate As Date, ByVal mPTax As Double) As Double
        '==============================================================================='
        ' Modified By : Aiby                                                            '
        '             : For                                        '
        '==============================================================================='
       
        Dim dtFromDt As Variant
        Dim mNoOfMonths As Long
        Dim mAmount     As Double
        Dim dtFromDate  As Date
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation Mode 1= Act and 2 = Circular                          '
        '-------------------------------------------------------------------------------'
        If gbFineCalculationMode = 1 Then
            If mPeriodID = 1 Then
                dtFromDt = DateSerial(mYearID, 10, 1)
            Else
                dtFromDt = DateSerial(mYearID + 1, 4, 1)
            End If
            
            If mYearID = gbFinancialYearID And mPeriodID = 2 Then
                Fine = 0
                Exit Function
            End If
            
            If mYearID < 2006 Then
                If mYearID = 2005 And mPeriodID = 2 Then
                    GoTo Skip
                End If
                If mUptoDate > DateSerial(2005, 9, 1) Then
                    'mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 4, 1), dtFromDt)) * 2 + 10
                    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 9, 1), dtFromDt)) * 2
                    mNoOfMonths = mNoOfMonths + 1
                    dtFromDt = DateSerial(2005, 10, 1)
                    mYearID = 2005
                    mPeriodID = 2
                Else
                    mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2
                    dtFromDt = mUptoDate
                    mYearID = Year(dtFromDt)
                    If Month(dtFromDt) > 9 And Month(dtFromDt) < 4 Then
                        mPeriodID = 2
                    Else
                        mPeriodID = 1
                    End If
                End If
                
            
                'If Year(mUptoDate) = 2005 Then
'                If mYearID = 2005 Then
'                    If mPeriodID = 1 Then
'                        mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
'                        If Month(mUptoDate) > 5 Then
'                            mNoOfMonths = mNoOfMonths - ((Month(mUptoDate) - 5) * 12)
'                        End If
'                    Else
'                        GoTo Skip:
'                    End If
                'End If
                'If Year(mUptoDate) < 2005 Then 'New Change For UptoDate
                'Else
                'If mYearID < 2005 Then 'New Change For UptoDate
                '    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 5, 1), dtFromDt)) * 2 + 10
                '    dtFromDt = DateSerial(2005, 11, 1)
                'End If
                'Else
                 '   mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
                'End If
                
                
                
            End If
Skip:
            If mUptoDate >= dtFromDt Then
                'mNoOfMonths = mNoOfMonths + (gbFinancialYearID - mYearID) * 12 'New Change For UptoDate
                mNoOfMonths = mNoOfMonths + 1 + Abs(DateDiff("M", mUptoDate, dtFromDt))  'New Change For UptoDate
            End If
            If mYearID = gbFinancialYearID And mPeriodID = 1 Then
                'mNoOfMonths = mNoOfMonths - 1
            End If
            'mNoOfMonths = mNoOfMonths + 1
            dtFromDate = DateAdd("m", 1, mUptoDate)
            'Debug.Print "No of Months (Fine) " & mNoOfMonths
            Fine = mPTax * mNoOfMonths / 100
            'If mNoOfMonths = 60 Then Stop
            Debug.Print "No of Months (Fine) " & mNoOfMonths & "    " & Fine
            Exit Function
        ElseIf gbFineCalculationMode = 2 Then
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation As Per Circular                                       '
        '-------------------------------------------------------------------------------'
           'mPTax = Format(mPTax * 2, "0.00")
            dtFromDt = DateSerial(mYearID, 11, 1)
            If mYearID = gbFinancialYearID Then
                Fine = 0
                Exit Function
            End If
            If mYearID < 2005 Then
                mNoOfMonths = Abs(DateDiff("m", DateSerial(2005, 8, 1), dtFromDt))
                dtFromDt = DateSerial(2005, 9, 1)
                mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt))
            End If
            mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt)) + 1
            Fine = mPTax * mNoOfMonths / 100
            Exit Function
        End If
    End Function
    Private Sub CalculateDemand()
        Dim mPeriodID As Integer
        Dim mNoOfDemands As Integer
        Dim mYear1 As Integer
        Dim mYear2 As Integer
        mYear1 = val(txtFromYear)
        mYear2 = val(txtToYear)
        mPeriodID = val(txtFromPeriodID)
        If mPeriodID > 1 Then mPeriodID = 2 Else mPeriodID = 1
        mNoOfDemands = mYear2 - mYear1 + 1
        If optHalfYearRate.value Then
            mNoOfDemands = mNoOfDemands * 2
            If mPeriodID = 2 Then mNoOfDemands = mNoOfDemands - 1
            If val(txtToPeriodID) = 1 Then mNoOfDemands = mNoOfDemands - 1
        End If
        lblMsg.Caption = "No of Half Years : " & mNoOfDemands
        mNoOfHalfYears = mNoOfDemands
    End Sub
    Private Sub CalculateWithLC()
        Dim mAmt As Double
        Dim mLC As Double
        Dim mPT As Double
        
        mAmt = val(txtTaxRate)
        mLC = (mAmt * 2 * 5 / 100) / 2
        If mLC - Int(mLC) > 0 Then
            mLC = mLC + (1 - (mLC - Int(mLC)))
        End If
        'lblLC.Caption = Format(mAmt * (5 / 100), "0.00")
        lblLC.Caption = Format(mLC, "#0")
        lblPT.Caption = Format(mAmt, "0.00")
    End Sub
    Private Sub Calculate()
        Dim mAmt As Double
        mAmt = val(txtTaxRate)
        lblLC.Caption = Format(mAmt / 21, "0.00")
        lblPT.Caption = Format(mAmt - (mAmt / 21), "0.00")
    End Sub
    Private Sub FormInitialize()
        txtTaxRate.Text = ""
        optHalfYearRate.value = True
        txtFromYear.Text = ""
        txtToYear.Text = ""
        txtFromPeriodID.Text = ""
        txtToPeriodID.Text = ""
        lblMsg.Caption = ""
        mHYrTaxInFraction = False
    End Sub

    Private Sub chkFineWaiver_Click()
        If chkFineWaiver.value = 1 Then
            frmFineWaiver.Mode = 2
            frmFineWaiver.Show vbModal, frmPropertyTax
        End If
    End Sub

Private Sub chkNonR_Click()
 If chkNonR.value = vbChecked Then
    NonResi = 1
 Else
    NonResi = 0
 End If
    
End Sub

    Private Sub chkSTax_Click()
           If chkSTax.value = 1 Then
               Call CopyToReciept
               frmPTaxCalculator.Mode = 5
               frmPTaxCalculator.Show vbModal
            End If
    End Sub

'    Private Sub chkSTax_GotFocus()
'        If txtTaxRate.Text <> "" And txtFromYear.Text <> "" And txtFromPeriodID.Text <> "" And txtToYear.Text <> "" And txtToPeriodID.Text <> "" Then
'            chkSTax.value = vbChecked
'        Else
'            MsgBox ("Please Enter Values")
'            Exit Sub
'            chkSTax.value = vbUnchecked
'            txtTaxRate.SetFocus
'        End If
'    End Sub


    Private Sub cmdCalculate_Click()
        Dim mFine As Double
        Dim mLoop As Long
        Dim mYearID As Integer
        Dim mPeriodID As Integer
        
        
        If val(txtTaxRate) = 0 Then
            MsgBox "Enter the TaxRate", vbInformation
            txtTaxRate.SetFocus
            Exit Sub
        End If
        If val(txtFromYear) = 0 Then
            MsgBox "Enter the Year", vbInformation
            txtFromYear.SetFocus
            Exit Sub
        End If
        If val(txtFromPeriodID) = 0 Then
            MsgBox "Enter the Period", vbInformation
            txtFromPeriodID.SetFocus
            Exit Sub
        End If
        
        If val(txtToYear) = 0 Then
            MsgBox "Enter the Year", vbInformation
            txtToYear.SetFocus
            Exit Sub
        End If
        If val(txtToPeriodID) = 0 Then
            MsgBox "Enter the Period", vbInformation
            txtToPeriodID.SetFocus
            Exit Sub
        End If
        
       ' chkSTax.Visible = True
        Call CalculateDemand
        mYearID = val(txtFromYear)
        mPeriodID = val(txtFromPeriodID)
        mFine = 0
        Dim dtToDate As Date
        If val(txtToPeriodID) = 2 Then
            dtToDate = (DateSerial(val(txtToYear.Text) + 1, 3, 28))
        Else
            dtToDate = (DateSerial(val(txtToYear.Text), 9, 28))
        End If
        If dtToDate > gbTransactionDate Then
            dtToDate = gbTransactionDate
        End If
       ' mFine = Fine(mYearID, mPeriodID, dtToDate, Val(txtTaxRate.Text))
        Dim mLoopFlag As Boolean
        mLoopFlag = True
        While mLoopFlag 'And (mYearID <= val(txtToYear) And mPeriodID <= val(txtToPeriodID))
            mFine = mFine + Fine(mYearID, mPeriodID, gbTransactionDate, val(txtTaxRate.Text))
            'If mYearID = val(txtToYear) And mPeriodID = val(txtToPeriodID) Then ' Changed on 26-07-10 to check 2000/2
            If mYearID >= val(txtToYear) And mPeriodID = val(txtToPeriodID) Then
                mLoopFlag = False
            Else
                If mPeriodID < 2 Then
                    mPeriodID = 2
                Else
                    mYearID = mYearID + 1
                    mPeriodID = 1
                End If
            End If
            If mYearID > val(txtToYear) Then
                'If mPeriodID > val(txtToPeriodID) Then ' Changed by Aiby 26-07-10 Check Runtime Error 2000/2
                If mPeriodID >= val(txtToPeriodID) Then
                    mLoopFlag = False
                End If
            End If
        Wend
        'For mLoop = 1 To mNoOfHalfYears
            'mFine = mFine + CalculateFineforPTax(mYearID, mPeriodID, Val(lblPT.Caption))
        'Next mLoop
        
        If mFine - Int(mFine) > 0 Then
            mFine = mFine + (1 - (mFine - Int(mFine)))
        End If
        lblFine.Caption = Format(mFine, "#0")
        cmdCopyToReceipt.SetFocus
        
    End Sub
    Private Sub cmdCopyToReceipt_Click()
    
    Call CopyToReciept
'        Dim mLoop As Integer
'        Dim mYearID As Integer
'        Dim mPeriodID As Integer
'        Dim objAcc As New clsAccounts
'        Dim mRow As Integer
'        Dim mFineFlag As Boolean
'        Dim mAmt As Double
'
'
'        If chkSTax.value = 1 Then
'            mRowCount = mRowCount + 1
'        End If
'        If chkSTax.value = 0 Then
'            mRowCount = 1
'        End If
'
'                   If mDemandMode = True Then
'                    copyToDemand
'                    Exit Sub
'                End If
'                mYearID = val(txtFromYear)
'                mPeriodID = val(txtFromPeriodID)
'                mRow = mRowCount
''                frmReceiptsCounter.vsGrid.Clear 1, 0
'                For mLoop = 1 To mNoOfHalfYears
'                    If mRow > 8 Then
'                        'MsgBox "Can't Print more than 8 Rows...", vbInformation
'                        'Exit For
'                    End If
'                    If mYearID < gbFinancialYearID Then
'                        mFineFlag = True
'                        objAcc.SetAccountCode gbAcHeadCodePropertyTaxArrear
'                    Else
'                        objAcc.SetAccountCode gbAcHeadCodePropertyTaxCurrent
'                    End If
'                    If frmReceiptsCounter.vsGrid.Rows = mRow Then
'                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 10
'                    End If
'
'                    If objAcc.AccountHeadID > 0 Then
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
'
'                        If mPeriodID = 1 Then
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 1 '"First Half"
'                        Else
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 2 '"Second Half"
'                        End If
'                        If mHYrTaxInFraction Then
'                            If mPeriodID = 1 Then
'                                mAmt = val(lblPT.Caption) + 0.5
'                            Else
'                                mAmt = val(lblPT.Caption) - 0.5
'                            End If
'                        Else
'                            mAmt = Format(val(lblPT.Caption), "#0")
'                        End If
'
'                        If mYearID < gbFinancialYearID Then
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = mAmt
'                        Else
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = mAmt
'                        End If
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
'
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = mAmt
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
'                    End If
'
'                    mRow = mRow + 1
'                    If frmReceiptsCounter.vsGrid.Rows = mRow Then
'                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
'                    End If
'
'                    objAcc.SetAccountCode gbAcHeadCodeLibraryCess
'                    If objAcc.AccountHeadID > 0 Then
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
'                        If mPeriodID = 1 Then
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "First Half"
'                        Else
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "Second Half"
'                        End If
'                        If mYearID < gbFinancialYearID Then
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = val(lblLC.Caption)
'                        Else
'                            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblLC.Caption)
'                        End If
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
'
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblLC.Caption)
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
'                    End If
'
'                    mRow = mRow + 1
'                    If frmReceiptsCounter.vsGrid.Rows = mRow Then
'                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
'                    End If
'
'                    If mPeriodID = 1 Then
'                        mPeriodID = 2
'                    Else
'                        mPeriodID = 1
'                        mYearID = mYearID + 1
'                    End If
'                Next mLoop
'                mRowCount = mRow
'                'If mFineFlag Then
'                If val(lblFine.Caption) > 0 Then
'                    objAcc.SetAccountCode gbAcHeadCodePenalInterest
'                    If objAcc.AccountHeadID > 0 Then
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = ""
'
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblFine.Caption)
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = gbFinancialYearID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblFine.Caption)
'                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
'                    End If
'                End If
'                frmReceiptsCounter.Calculate
'                frmReceiptsCounter.txtTransactionType.Tag = gbTransactionTypePTax
'                frmReceiptsCounter.txtTransactionType.Text = "Property Tax"
'                Unload Me
'
    End Sub
    Private Sub Form_Activate()
        Me.Left = (frmMenu.Width - Me.Width) / 2
        Me.Top = 2500
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
        End If
    End Sub
Private Sub Form_Load()
    If (gbLBPanchayat) Then
        chkNonR.Visible = True
        frmReceiptsCounter.vsGrid.Clear 1, 0
    End If
   ' chkSTax.Visible = False
    NonResi = 0
End Sub

    Private Sub optFullYearRate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub optHalfYearRate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub txtFromPeriodID_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtFromPeriodID_LostFocus()
        If Trim(txtFromPeriodID.Text) = "" Then
            txtFromPeriodID.Text = 1
        End If
    End Sub
    Private Sub txtFromYear_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtFromYear_LostFocus()
        Dim mYear As Integer
        mYear = val(txtFromYear)
        If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
        If mYear < 1901 Then mYear = gbFinancialYearID
        txtFromYear = mYear
    End Sub
    Private Sub txtTaxRate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub txtTaxRate_LostFocus()
        Dim mTaxRate As Single
        mTaxRate = val(txtTaxRate.Text)
        
        If (mTaxRate - Int(mTaxRate)) > 0 Then
            mTaxRate = Int(mTaxRate) + 0.5
            mHYrTaxInFraction = True
        Else
            mHYrTaxInFraction = False
        End If
        txtTaxRate.Text = Format(mTaxRate, "0.00")
        'txtTaxRate.Text = Format(val(txtTaxRate), "#0")
        txtTaxRate.Text = Format(val(txtTaxRate), "0.00")
        
        Call CalculateWithLC
       
    End Sub
    Private Sub txtToPeriodID_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtToPeriodID_LostFocus()
        If Trim(txtToPeriodID) = "" Then
            txtToPeriodID.Text = "2"
        End If
    End Sub
    Private Sub txtToYear_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
    End Sub
    Private Sub txtToYear_LostFocus()
        Dim mYear As Integer
        mYear = val(txtToYear)
        If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
        If mYear < 1901 Then mYear = gbFinancialYearID
        txtToYear = mYear
    End Sub
    
    Public Property Let DemandMode(mVal As Boolean)
        mDemandMode = mVal
    End Property
    
    Public Property Get DemandMode() As Boolean
        DemandMode = mDemandMode
    End Property
    Public Property Let PTRowCount(mVal As Integer)
        mPTRow = mVal
    End Property
    
    Public Property Get PTRowCount() As Integer
        PTRowCount = mPTRow
    End Property
    Public Property Let NonResi(mVal As Integer)
        mNonR = mVal
    End Property
    Public Property Get NonResi() As Integer
        NonResi = mNonR
    End Property
    Public Sub CopyToReciept()
    Dim mLoop As Integer
    Dim mYearID As Integer
    Dim mPeriodID As Integer
    Dim objAcc As New clsAccounts
    Dim mRow As Integer
    Dim mFineFlag As Boolean
    Dim mAmt As Double

    CalculateDemand
     If txtTaxRate > 0 Then
             If NonResi = 1 Then
                        mRow = 1
                        If mDemandMode = True Then
                        copyToDemand
                        Exit Sub
                        End If
                        mYearID = val(txtFromYear)
                        mPeriodID = val(txtFromPeriodID)
                        frmReceiptsCounter.vsGrid.Clear 1, 0
                        For mLoop = 1 To mNoOfHalfYears
                        If mRow > 8 Then
                        End If
                        If mYearID < gbFinancialYearID Then
                        mFineFlag = True
                        objAcc.SetAccountCode gbAcHeadCodePropertyTax_NonResidential_Arrear
                        Else
                        objAcc.SetAccountCode gbAcHeadCodePropertyTax_NonResidential_Current
                        End If
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 10
                        End If
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                        
                        If mPeriodID = 1 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 1 '"First Half"
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 2 '"Second Half"
                        End If
                        If mHYrTaxInFraction Then
                        If mPeriodID = 1 Then
                        mAmt = val(lblPT.Caption) + 0.5
                        Else
                        mAmt = val(lblPT.Caption) - 0.5
                        End If
                        Else
                        mAmt = Format(val(lblPT.Caption), "#0")
                        End If
                        
                        If mYearID < gbFinancialYearID Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = mAmt
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = mAmt
                        End If
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = mAmt
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        End If
                        
                        mRow = mRow + 1
                        PTRowCount = mRow
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
                        End If
                        
                        objAcc.SetAccountCode gbAcHeadCodeLibraryCess
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                        If mPeriodID = 1 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "First Half"
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "Second Half"
                        End If
                        If mYearID < gbFinancialYearID Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = val(lblLC.Caption)
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblLC.Caption)
                        End If
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblLC.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        End If
                        
                        mRow = mRow + 1
                        PTRowCount = mRow
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
                        End If
                        
                        If mPeriodID = 1 Then
                        mPeriodID = 2
                        Else
                        mPeriodID = 1
                        mYearID = mYearID + 1
                        End If
                        Next mLoop
                        ' mRowCount = mRow
                        PTRowCount = mRow
                        'If mFineFlag Then
                        If val(lblFine.Caption) > 0 Then
                        objAcc.SetAccountCode gbAcHeadCodePenalInterest
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblFine.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = gbFinancialYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblFine.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        
                        End If
                        PTRowCount = mRow + 1
                        End If
                        
                Else
                
                          
                        mRow = 1
                        If mDemandMode = True Then
                        copyToDemand
                        Exit Sub
                        End If
                        mYearID = val(txtFromYear)
                        If mYearID = 0 Then
                            mYearID = gbFinancialYearID
                        End If
                        mPeriodID = val(txtFromPeriodID)
                        If mPeriodID = 0 Then
                            mPeriodID = 1
                        End If
                        '              mRow = mRowCount
                        frmReceiptsCounter.vsGrid.Clear 1, 0
                        For mLoop = 1 To mNoOfHalfYears
                        If mRow > 8 Then
                        
                        End If
                        If mYearID < gbFinancialYearID Then
                        mFineFlag = True
                        objAcc.SetAccountCode gbAcHeadCodePropertyTaxArrear
                        Else
                        objAcc.SetAccountCode gbAcHeadCodePropertyTaxCurrent
                        End If
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 10
                        End If
                        
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                        
                        If mPeriodID = 1 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 1 '"First Half"
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 2 '"Second Half"
                        End If
                        If mHYrTaxInFraction Then
                        If mPeriodID = 1 Then
                            mAmt = val(lblPT.Caption) + 0.5
                        Else
                            mAmt = val(lblPT.Caption) - 0.5
                        End If
                        Else
                        mAmt = Format(val(lblPT.Caption), "#0")
                        End If
                        
                        If mYearID < gbFinancialYearID Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = mAmt
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = mAmt
                        End If
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = mAmt
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        End If
                        
                        mRow = mRow + 1
                        PTRowCount = mRow
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
                        End If
                        
                        objAcc.SetAccountCode gbAcHeadCodeLibraryCess
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                        If mPeriodID = 1 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "First Half"
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = "Second Half"
                        End If
                        If mYearID < gbFinancialYearID Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = val(lblLC.Caption)
                        Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblLC.Caption)
                        End If
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblLC.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        End If
                        
                        mRow = mRow + 1
                        PTRowCount = mRow
                        If frmReceiptsCounter.vsGrid.Rows = mRow Then
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
                        End If
                        
                        If mPeriodID = 1 Then
                        mPeriodID = 2
                        Else
                        mPeriodID = 1
                        mYearID = mYearID + 1
                        End If
                        Next mLoop
                        ' mRowCount = mRow
                        PTRowCount = mRow
                        'If mFineFlag Then
                        If val(lblFine.Caption) > 0 Then
                        objAcc.SetAccountCode gbAcHeadCodePenalInterest
                        If objAcc.AccountHeadID > 0 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = ""
                        
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = val(lblFine.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = gbFinancialYearID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodID
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(lblFine.Caption)
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                        
                        End If
                        PTRowCount = mRow + 1
                        End If
        
                End If

     Else
        frmPropertyTaxCalculator.PTRowCount = 1
     End If

    frmReceiptsCounter.Calculate
    frmReceiptsCounter.txtTransactionType.Tag = gbTransactionTypePTax
    frmReceiptsCounter.txtTransactionType.Text = "Property Tax"
    frmPTaxCalculator.FormInitialization

    Unload Me
                      
End Sub
