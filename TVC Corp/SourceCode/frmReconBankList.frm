VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmReconBankList 
   BackColor       =   &H00FEFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconciliation Bank List"
   ClientHeight    =   9420
   ClientLeft      =   -360
   ClientTop       =   -1560
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15210
      TabIndex        =   17
      Top             =   0
      Width           =   15240
      Begin VB.Image imgNext 
         Height          =   420
         Left            =   14610
         Picture         =   "frmReconBankList.frx":0000
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANK RECONCILIATION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   255
         TabIndex        =   18
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmBankDetails 
      BackColor       =   &H00FEFFFF&
      Height          =   8040
      Left            =   8040
      TabIndex        =   9
      Top             =   1200
      Width           =   7035
      Begin VB.CommandButton cmdReport 
         Caption         =   "REPORT"
         Height          =   450
         Left            =   6150
         TabIndex        =   27
         Top             =   6720
         Width           =   855
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6150
         TabIndex        =   24
         Top             =   7305
         Width           =   855
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6015
         TabIndex        =   3
         Top             =   2565
         Width           =   795
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   0
         ScaleHeight     =   645
         ScaleWidth      =   6990
         TabIndex        =   22
         Top             =   0
         Width           =   7020
         Begin VB.Label lblBankDetails 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANK DETAILS"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   150
            TabIndex        =   23
            Top             =   165
            Width           =   1620
         End
      End
      Begin VB.TextBox txtBankStatementAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   3645
         TabIndex        =   2
         Top             =   2685
         Width           =   1710
      End
      Begin VB.TextBox txtBankBookAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2340
         Width           =   1710
      End
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   900
         Width           =   5610
      End
      Begin VB.TextBox txtHeadCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   1230
         Width           =   1335
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3030
         TabIndex        =   10
         Top             =   1245
         Width           =   3780
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridMonth 
         Height          =   4080
         Left            =   1185
         TabIndex        =   4
         Top             =   3270
         Width           =   4200
         _cx             =   7408
         _cy             =   7197
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   14271125
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15790320
         BackColorAlternate=   -2147483643
         GridColor       =   15987699
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReconBankList.frx":06C2
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   3480
         TabIndex        =   13
         Top             =   7275
         Width           =   1905
         Begin VB.CommandButton cmdYearDown 
            Caption         =   "<<"
            Height          =   345
            Left            =   30
            TabIndex        =   6
            Top             =   143
            Width           =   525
         End
         Begin VB.CommandButton cmdYearUp 
            Caption         =   ">>"
            Height          =   345
            Left            =   1305
            TabIndex        =   7
            Top             =   143
            Width           =   525
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "####-##"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   600
            TabIndex        =   5
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.Label lblLastDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2670
         TabIndex        =   25
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCES AS ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   21
         Top             =   1935
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   1185
         X2              =   5355
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line1 
         X1              =   1185
         X2              =   5355
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblBankStatementAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AS PER BANK STATEMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1305
         TabIndex        =   1
         Top             =   2700
         Width           =   2070
      End
      Begin VB.Label lblBankBookAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AS PER BANK BOOK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1290
         TabIndex        =   20
         Top             =   2355
         Width           =   1545
      End
      Begin VB.Label lblBankName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   930
         Width           =   435
      End
      Begin VB.Label lblHeadCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HEAD CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   255
         TabIndex        =   15
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2595
         TabIndex        =   14
         Top             =   1275
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FAFDFD&
      Height          =   7950
      Left            =   195
      ScaleHeight     =   7890
      ScaleMode       =   0  'User
      ScaleWidth      =   7530
      TabIndex        =   8
      Top             =   1245
      Width           =   7590
      Begin VB.CommandButton cmdClickBankListGrid 
         Caption         =   "ACTIVATE DBL CLICK EVENT"
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
         Left            =   225
         TabIndex        =   26
         Top             =   7470
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   7245
         Left            =   -30
         TabIndex        =   0
         Top             =   0
         Width           =   7575
         _cx             =   13361
         _cy             =   12779
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   14271125
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15790320
         BackColorAlternate=   -2147483643
         GridColor       =   15987699
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReconBankList.frx":072B
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
End
Attribute VB_Name = "frmReconBankList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mdtStartDate As Variant             ' FIRST DAY OF SELECTED OR NEXT RECONCILIATION MONTH
Dim mdtLastDate As Variant              ' LAST DAY OF SELECTED MONTH
Dim mReconciliationStarted As Boolean   ' A CONTROL VARIABLE TO CHECK WHETHER RECONCILIATION STARTED FOR THE SELECTED BANK OR NOT
Dim mEditFlag As Boolean                ' A CoNTROL VARIABLE TO CHECK EDIT MODE
Dim mReconStatus As Integer             ' A CONTROL VARIABLE TO STORE RECONCILIATION's CURRENT STATUS
Dim mFirstTimeReconFlag As Boolean
Dim mNextDate As Variant                ' NEXT RECONCILIATION DATE to SET
Dim mResetFlag  As Integer              'Reconciliation Reset flag mResetFlag=1 --Reset completely mResetFlag=2 Month wise reset from last month

Private Sub FormInitialize()
    txtHeadCode.Tag = ""                ' BANKE ACCOUNT HEAD ID STORING
    txtHeadCode.Text = ""
    txtBankName.Text = ""
    txtBankName.Tag = ""
    txtBankBookAmount.Text = ""
    txtBankBookAmount.ToolTipText = ""
    txtBankBookAmount.BackColor = &HE0E0E0
    
    txtBankStatementAmount.Text = ""
    txtBankStatementAmount.Tag = ""     ' RECONCILIATION ID STORING
    txtType.Text = ""
    txtType.Tag = ""                    ' MINIOR ACCOUNT HEAD ID STORES
    lblLastDate.Caption = "#"
    mdtStartDate = Null
    mdtLastDate = Null
    mReconciliationStarted = False
    cmdStart.Caption = "START"
    mEditFlag = False
    mReconStatus = -1
    mFirstTimeReconFlag = False
    
    lblYear.Caption = Trim(str(gbFinancialYearID)) + "-" + Trim(str((gbFinancialYearID - 2000)))
    lblYear.Tag = gbFinancialYearID - 1
    'cmdReset.Visible = False
    
End Sub
Private Sub CheckReconciled()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        
        If val(txtHeadCode.Tag) > 0 Then
            mSql = "SELECT *  FROM faBankReconcile "
            mSql = mSql + " Where intBankAccountHeadID = " & val(txtHeadCode.Tag)
            mSql = mSql + " AND  intYearID = " & (lblYear.Tag) & " And intMonthID = " & val(vsGridMonth.TextMatrix(vsGridMonth.Row, 0))
            objDb.SetConnection mCnn
            Set Rec = mCnn.Execute(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                txtBankBookAmount.Text = Format(Rec!numBankBookBalance, "0.00")
                txtBankStatementAmount.Text = Format(Rec!numPassBookBalance, "0.00")
                txtBankStatementAmount.Tag = Rec!intReconID
                If IsNumeric(Rec!tnyReconStatus) Then
                    If Rec!tnyReconStatus = 1 Then
                        cmdStart.Enabled = False
                    End If
                End If
                mReconStatus = Rec!tnyReconStatus
            Else
                
            End If
        End If
    End Sub
    Private Sub ClearBalanceFields()
        lblLastDate.Caption = ""
        txtBankBookAmount.Text = ""
        txtBankStatementAmount.Text = ""
        txtBankStatementAmount.Tag = ""
        cmdStart.Caption = "START"
    End Sub
    
    
    Private Sub GetBankBookBalance()
        Dim objLdgr As New clsAccounts
        Dim mAmt As Double
    
        txtBankBookAmount.Text = ""
        If val(txtHeadCode.Tag) > 0 Then
            If IsDate(mdtLastDate) Then
                mAmt = objLdgr.GetLedgerBalance(val(txtHeadCode.Tag), mdtLastDate)
                txtBankBookAmount.Text = Format(mAmt, "0.00   ")
                txtBankStatementAmount.Text = ""
                txtBankStatementAmount.Enabled = True
                txtBankStatementAmount.SetFocus
            End If
        End If
    End Sub

Private Sub SetBalance(mDate As Date)
    'PRECONDITION : MUST SET THE BANK
    If val(txtHeadCode.Tag) = 0 Then
        Exit Sub
    End If
    
    
    'VALIDATING DATE WITH RECONSTART DATE LAST RECOND DATE
    Dim objBank As New clsBank
    Dim objLdgr As New clsAccounts
    Dim mBankBalance As Double
    objBank.SetBankInfoByAccID (val(txtHeadCode.Tag))
    If objBank.BankID > 0 Then
        If mDate < IIf(IsDate(objBank.ReconciliationStartDate), objBank.ReconciliationStartDate, mDate) Then
            mDate = CDate(objBank.ReconciliationStartDate)
        ElseIf mDate > IIf(IsDate(objBank.ReconciliationLastDate), objBank.ReconciliationLastDate, mDate) Then
            If IsDate(mNextDate) Then
                If mDate <> mNextDate Then
                    mDate = FindNextDate(CDate(objBank.ReconciliationLastDate))
                End If
            Else
                mDate = FindNextDate(CDate(objBank.ReconciliationLastDate))
            End If
        End If
        SetYear (mDate) 'SETS THE YEAR : ID and Caption
        mBankBalance = objLdgr.GetLedgerBalance(objBank.BankAccountHeadID, mDate)
        lblLastDate.Caption = DdMmmYy(mDate)
        mdtLastDate = DdMmmYy(mDate)
        
    Else
        Exit Sub ' BANK NOT SET
    End If
    
    'FIND ANY RECOND IN BANK RECONILIATION TABLE
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    Dim mCn As New ADODB.Connection
    Dim objDb As New clsDB
    Dim mYearID As Integer
    Dim mMonthID As Integer
    mYearID = Year(mDate)
    mMonthID = Month(mDate)
    If mMonthID < 4 Then
        mYearID = mYearID - 1
    End If
        
    txtBankBookAmount.Text = ""
    txtBankStatementAmount.Text = ""
    txtBankBookAmount.BackColor = vbWindowBackground
    cmdStart.Enabled = True
    mSql = "SELECT * FROM faBankReconcile WHERE intBankAccountHeadID = " & objBank.BankAccountHeadID
    mSql = mSql + " AND intYearID = " & mYearID & " AND intMonthID = " & mMonthID
    objDb.SetConnection mCn
    Rec.Open mSql, mCn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
        txtBankStatementAmount.Tag = Rec!intReconID
        txtBankBookAmount.Text = Format(Rec!numBankBookBalance, "0.00")
        txtBankStatementAmount.Text = Format(Rec!numPassBookBalance, "0.00")
        txtBankStatementAmount.Enabled = False
        cmdStart.Caption = "EDIT"
        If Rec!tnyReconStatus = 1 Then
            cmdStart.Enabled = False
        End If
        If val(txtBankBookAmount) <> mBankBalance Then
            txtBankBookAmount.BackColor = &H9898FF
            txtBankBookAmount.ToolTipText = "Bank Balance is changed after starting RECONCILIATION"
        End If
        mReconStatus = Rec!tnyReconStatus
    Else
        cmdStart.Caption = "START"
        txtBankStatementAmount.Tag = ""
        txtBankBookAmount.Text = Format(mBankBalance, "0.00")
        txtBankStatementAmount.Text = ""
        txtBankStatementAmount.Enabled = True
        txtBankStatementAmount.SetFocus
        mReconStatus = -1
    End If
    Rec.Close
    
    
End Sub
Private Sub SetYear(mDt As Date)
    
        Dim mYearID As Integer
        Dim mMonthID As Integer
        mYearID = Year(mDt)
        mMonthID = Month(mDt)
        If mMonthID < 4 Then
            mYearID = mYearID - 1
        End If
        lblYear.Caption = str(mYearID) & "-" & Right(str(mYearID + 1), 2)
        lblYear.Tag = mYearID
        If mYearID <> gbFinancialYearID Then
            lblYear.ForeColor = &HFF&
        Else
            lblYear.ForeColor = vbDefault
        End If
    
End Sub

Private Function FindNextDate(mDt As Date) As Date
    mDt = DateSerial(Year(mDt), Month(mDt), 1)
    mDt = DateAdd("m", 2, mDt)
    mDt = DateAdd("d", -1, mDt)
    FindNextDate = mDt
End Function
Private Function FindLastDateOfTheMonth(mDt As Date) As Date
    
    mDt = DateSerial(Year(mDt), Month(mDt), 1)
    mDt = DateAdd("m", 1, mDt)
    mDt = DateAdd("d", -1, mDt)
    FindLastDateOfTheMonth = mDt
End Function
Private Sub UpdateMonthlyStatus(mFinYearID As Integer)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim objDb As New clsDB
    Dim mLoop As Integer
    Dim mRecCount As Integer
    For mLoop = 1 To 12
        vsGridMonth.Cell(flexcpBackColor, mLoop, 0, , 2) = &H80000005
    Next
    
    'UPDATING STATUS
    mSql = " SELECT * FROM faBankReconcile WHERE intBankAccountHeadID = " & val(txtHeadCode.Tag)
    mSql = mSql + " AND intYearID = " & mFinYearID
    mSql = mSql + " ORDER by (intYearID + intMonthID)"
    objDb.SetConnection mCnn
    Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
        While Not Rec.EOF
            For mLoop = 1 To 12
                If vsGridMonth.TextMatrix(mLoop, 0) = Rec!intMonthID Then
                    If IIf(IsNull(Rec!tnyReconStatus), 0, Rec!tnyReconStatus) = 1 Then
                         vsGridMonth.Cell(flexcpBackColor, mLoop, 0, , 2) = &HB2FFBD ' Green
                         If mLoop < 12 Then
                            If vsGridMonth.Cell(flexcpBackColor, mLoop + 1, 0, , 2) <> &HB2FFBD Then
                                vsGridMonth.Cell(flexcpBackColor, mLoop + 1, 0, , 2) = &HAFCCCC
                            End If
                         End If
                    Else
                        vsGridMonth.Cell(flexcpBackColor, mLoop, 0, , 2) = &HAFCCCC
                        Call SetBalance(CDate(mNextDate))
                    End If
                End If
            Next
            Rec.MoveNext
        Wend
    Else
    
        Rec.Close
        mSql = " SELECT * FROM faBankReconcile WHERE intBankAccountHeadID = " & val(txtHeadCode.Tag)
        mSql = mSql + " AND intMonthID=3 AND intYearID = " & mFinYearID - 1
        mSql = mSql + " ORDER by (intYearID + intMonthID)"
        objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rec.BOF And Rec.EOF) Then
            vsGridMonth.Cell(flexcpBackColor, 1, 0, , 2) = &HAFCCCC
            Call SetBalance(CDate(mNextDate))
        End If
    End If
    Rec.Close
    vsGridMonth.HighLight = flexHighlightNever
    
    '''    ' NOTE: CHECKING WHETHER USER IS ALLOWED THE RESET THE STARTUP MONTH SELECTED
    '''    '     : Only allowed if no child records are updated agains the selected bank
    '''    '     : And count of parent table record will be 1 that time, status is not checking
    '''
    '''    mSQL = " Select Count(faBankReconcile.intReconID) ParentCount, Count( faBankReconcileChild.intReconID) ChildCount From faBankReconcile"
    '''    mSQL = mSQL + " LEFT JOIN faBankReconcileChild ON faBankReconcileChild.intReconID = faBankReconcile.intReconID"
    '''    mSQL = mSQL + " Where intBankAccountHeadID = " & val(txtHeadCode.Tag)
    '''    Rec.Open mSQL, mCnn, adOpenStatic, adLockReadOnly, adCmdText
    '''    If Not (Rec.BOF And Rec.EOF) Then
    '''        If Rec!ParentCount = 1 And Rec!ChildCount = 0 Then
    '''            cmdReset.Enabled = True
    '''        Else
    '''            cmdReset.Enabled = False
    '''        End If
    '''    End If
    '''    Rec.Close
    
End Sub

Private Sub cmdClickBankListGrid_Click()
    Call UpdateMonthlyStatus(val(lblYear.Tag))
    Dim mDt As Date
    mDt = FindNextDate(CDate(mNextDate))
    Call SetBalance(CDate(mNextDate))
End Sub

    Private Sub cmdReport_Click()
        Dim mMonth  As Integer
        Dim mYear   As Integer
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        
        If lblLastDate.Caption <> "#" Then
            If val(txtHeadCode.Tag) < 1 Then
                MsgBox "Please Select A Bank", vbApplicationModal
                Exit Sub
            End If
            If lblLastDate.Caption = "" Then
                MsgBox "Please Select A Month", vbApplicationModal
                Exit Sub
            End If
            mMonth = Month(lblLastDate.Caption)
            If mMonth < 4 Then
                mYear = Year(lblLastDate.Caption) - 1
            Else
                mYear = Year(lblLastDate.Caption)
            End If
            
            arInput = Array(val(txtHeadCode.Tag), mMonth, val(mYear))
            frmNewRpt.rptFileName = App.Path & "\Reports\rptReconciliation.rpt"
            frmNewRpt.WindowState = vbMaximized
            frmNewRpt.InputParameters = arInput
            Call frmNewRpt.ShowReport
            frmNewRpt.Show
        End If
    End Sub

 Private Function ResetValidation() As Boolean
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        ResetValidation = False
        If val(txtHeadCode.Tag) = 0 Then
            MsgBox "Please select a Bank/Treasury", vbInformation
            ResetValidation = False
            Exit Function
        End If
        objDb.SetConnection mCnn
        Rec.Open "Select * From faBanks Where intAccountHeadID=" & val(txtHeadCode.Tag), mCnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rec.BOF And Rec.EOF) Then
            If IsNull(Rec!dtReconStartDate) = True Then
                MsgBox "Reconciliation of Selected Bank not Started", vbInformation
                ResetValidation = False
                Exit Function
            End If
        End If
        
        ResetValidation = True
    End Function
    
'''Private Sub cmdReset_Click()
'''    Dim mSql As String
'''    Dim Rec As New ADODB.Recordset
'''    Dim mCnn As New ADODB.Connection
'''    Dim objDb As New clsDB
'''    Dim mRecCount As Integer
'''    Dim mReconID As Integer
'''
'''
'''    objDb.SetConnection mCnn
'''            If mFirstTimeReconFlag And mReconStatus < 1 Then
'''
'''                mReconID = val(txtBankStatementAmount.Tag)
'''                mSql = "SELECT * FROM faBankReconcile WHERE intReconID = " & mReconID
'''                Rec.CursorLocation = adUseClient
'''                Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
'''                If Not (Rec.BOF And Rec.EOF) Then
'''                    mRecCount = Rec.RecordCount
'''                    If mRecCount > 1 Or IIf(IsNull(Rec!tnyReconStatus), 0, Rec!tnyReconStatus) = 1 Then
'''                        mSql = "Previous Month(s) are already Reconciled" & vbCrLf
'''                        mSql = "There for, Application can not reset the start up Month" & vbCrLf
'''                        MsgBox mSql, vbInformation
''''                        cmdReset.Visible = False
'''                        Exit Sub
'''                    End If
'''                End If
'''                Rec.Close
'''
'''                'GET CONFIRMATION
'''                mSql = "This Will Reset the selected MONTH to start Reconciliation" & vbCrLf
'''                mSql = mSql + " Do you want to REST Now ?" & vbCrLf
'''                If MsgBox(mSql, vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'''                    Exit Sub
'''                End If
'''
'''                'DELETE AND REST DATES IN BANKs TABLE
'''                mSql = "Delete From faBankReconcileChild WHERE intReconID = " & mReconID
'''                mCnn.Execute mSql
'''
'''                mSql = "Delete From faBankReconcile WHERE intReconID = " & mReconID
'''                mCnn.Execute mSql
'''
'''                mSql = "Update faBanks Set dtReconStartDate = Null, dtReconEndDate = Null WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
'''                mCnn.Execute mSql
'''
'''                cmdReset.Visible = False
'''                Call FormInitialize
'''            End If
'''End Sub


Private Sub cmdReset_Click()
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    Dim Rec1 As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objDb As New clsDB
    Dim mRecCount As Integer
    Dim mReconID As Integer
    Dim mLastMonth As Integer
    Dim mLastYear As Integer
    Dim mLastYearForMsg As Integer
    Dim mLastDateUpdate   As Date
    Dim mStartDate As Date
    Dim mEndDate As Date
    objDb.SetConnection mCnn
    If ResetValidation = True Then
        mSql = " Select intYearID,intMonthID,intReconID From faBankReconcile Where intBankAccountHeadID=" & val(txtHeadCode.Tag)
        mSql = mSql + " AND intReconID=(Select Max(intReconID) From faBankReconcile Where intBankAccountHeadID=" & val(txtHeadCode.Tag) & ")"
        Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rec.BOF And Rec.EOF) Then
            mLastMonth = Rec!intMonthID
            mLastYear = Rec!intYearID
            mReconID = Rec!intReconID
            
            If mLastMonth < 4 Then
                mLastYearForMsg = mLastYear + 1
            Else
                mLastYearForMsg = mLastYear
            End If
            If MsgBox("Selected Bank Done Reconciliation Upto " & MonthName(mLastMonth) & " " & mLastYearForMsg & " Do you want to Proceed ", vbYesNo) = vbYes Then
'            If MsgBox("Selected Bank Done Reconciliation Upto " & MonthName(mLastMonth) & " " & mLastYearForMsg & " Do you want to RESET Completely", vbYesNo) = vbYes Then
'                If MsgBox("This Will Reset Your Reconciliation Completely For the selected Bank, Are you Sure....", vbCritical) = vbOK Then
'                    mSql = "Delete From faBankReconcileChild WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
'                    mCnn.Execute mSql
'
'                    mSql = "Delete From faBankReconcile WHERE intBankAccountHeadID = " & val(txtHeadCode.Tag)
'                    mCnn.Execute mSql
'
'                    mSql = "Update faBanks Set dtReconEndDate =  Null,dtReconStartDate=Null WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
'                    mCnn.Execute mSql
'                    UpdateMonthlyStatus (val(lblYear.Tag))
'                End If
           ' Else
                If MsgBox("Are you sure to reset Reconciliation for the Month " & MonthName(mLastMonth) & " " & mLastYearForMsg, vbYesNo) = vbYes Then
                   'DELETE AND RESET DATES IN BANKs TABLE
                    mSql = "Delete From faBankReconcileChild WHERE intReconID = " & mReconID
                    mCnn.Execute mSql
                    
                    mSql = "Delete From faBankReconcile WHERE intReconID = " & mReconID
                    mCnn.Execute mSql
                    
'                    mSql = " Select intYearID,intMonthID,intReconID From faBankReconcile Where intBankAccountHeadID=" & val(txtHeadCode.Tag)
'                    mSql = mSql + " AND intReconID=(Select Max(intReconID) From faBankReconcile Where intBankAccountHeadID=" & val(txtHeadCode.Tag) & ")"
                    mSql = " SEt DATEFORMAT dmy"
                    mSql = mSql + " Select intYearID,intMonthID,intReconID,Cast(Convert(varchar(25),'1/'+convert(varchar(3),intMonthID)+'/'+Convert(varChar(4),case When intMonthID <4 then intYearID+1 else intYearID end ),103) as DATETIME) As FirstDay,"
                    mSql = mSql + " DateAdd(day,-1,Dateadd(Month,1,Cast(Convert(varchar(25),'1/'+convert(varchar(3),intMonthID)+'/'+Convert(varChar(4),case When intMonthID <4 then intYearID+1 else intYearID end ),103) as DATETIME))) LastDAte,"
                    mSql = mSql + " DateAdd(day,-1,Dateadd(Month,0,Cast(Convert(varchar(25),'1/'+convert(varchar(3),intMonthID)+'/'+Convert(varChar(4),case When intMonthID <4 then intYearID+1 else intYearID end ),103) as DATETIME))) dtReconEndDAte"
                    mSql = mSql + " From faBankReconcile Where intBankAccountHeadID = " & val(txtHeadCode.Tag)
                    mSql = mSql + " AND intReconID=(Select Max(intReconID) From faBankReconcile Where intBankAccountHeadID=" & val(txtHeadCode.Tag) & ")"
                    Rec1.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (Rec1.BOF And Rec1.EOF) Then
                    mLastMonth = Rec1!intMonthID
                    mLastYear = Rec1!intYearID
                        If mLastMonth < 4 Then
                            mLastYearForMsg = mLastYear + 1
                        Else
                            mLastYearForMsg = mLastYear
                        End If
                        
                       ' mNextDate = DdMmmYy(CDate("1/" & Rec1!intMonthID & "/" & mLastYearForMsg))
                         'mLastDateUpdate = DateAdd("d", -1, DateAdd("m", -1, mLastDateUpdate))
'                         mLastDateUpdate = DateAdd("m", 1, mNextDate)
'                         mLastDateUpdate = DateAdd("d", -1, mNextDate)
                        mLastDateUpdate = Rec1!dtReconEndDate
                        mSql = "Update faBanks Set dtReconEndDate =  '" & DdMmmYy(CDate(mLastDateUpdate)) & "' WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
                        mCnn.Execute mSql
                        mSql = "Update faBankReconcile Set tnyReconStatus=0 WHERE intReconID = " & val(Rec1!intReconID)
                        mCnn.Execute mSql
                        'mNextDate = DateAdd("m", 1, mLastDateUpdate)
                        mNextDate = Rec1!LastDate
                        SetBalance (mNextDate)
                        Rec1.Close
                    Else
                        mSql = "Update faBanks Set dtReconEndDate = Null WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
                        mCnn.Execute mSql
                    End If
                    
                    
                   
                    
                    UpdateMonthlyStatus (val(lblYear.Tag))
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
          
        Else
             mSql = "Select * From faBanks where intAccountHeadID=" & val(txtHeadCode.Tag)
             Rec1.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
             If Not (Rec1.BOF And Rec1.EOF) Then
                If (IsNull(Rec1!dtReconStartDate)) And (IsNull(Rec1!dtReconEndDate)) Then
                    MsgBox "Selected Bank not Started For Reconciliation", vbInformation
                    Exit Sub
                ElseIf (IsNull(Rec1!dtReconEndDate)) And (Not (IsNull(Rec1!dtReconStartDate))) Then
                    If MsgBox("Selected Bank Started on  " & Rec1!dtReconStartDate & "  Do you want to Reset ?", vbYesNo) = vbYes Then
                        mSql = "Update faBanks Set dtReconStartDate = Null WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
                        mCnn.Execute mSql
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
             Else
                MsgBox "Selected Bank not Started For Reconciliation", vbInformation
                Exit Sub
             End If
        End If
  End If
    
       
'        If mResetFlag = 1 Then
        
''            If mFirstTimeReconFlag And mReconStatus < 1 Then
''
''                mReconID = val(txtBankStatementAmount.Tag)
''                mSql = "SELECT * FROM faBankReconcile WHERE intReconID = " & mReconID
''                Rec.CursorLocation = adUseClient
''                Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
''                If Not (Rec.BOF And Rec.EOF) Then
''                    mRecCount = Rec.RecordCount
''                    If mRecCount > 1 Or IIf(IsNull(Rec!tnyReconStatus), 0, Rec!tnyReconStatus) = 1 Then
''                        mSql = "Previous Month(s) are already Reconciled" & vbCrLf
''                        mSql = "There for, Application can not reset the start up Month" & vbCrLf
''                        MsgBox mSql, vbInformation
'''                        cmdReset.Visible = False
''                        Exit Sub
''                    End If
''                End If
''                Rec.Close
''
''                'GET CONFIRMATION
''                mSql = "This Will Reset the selected Bank to start Reconciliation" & vbCrLf
''                mSql = mSql + " Do you want to RESET Now ?" & vbCrLf
''                If MsgBox(mSql, vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
''                    Exit Sub
''                End If
''
''                'DELETE AND REST DATES IN BANKs TABLE
''                mSql = "Delete From faBankReconcileChild WHERE intReconID = " & mReconID
''                mCnn.Execute mSql
''
''                mSql = "Delete From faBankReconcile WHERE intReconID = " & mReconID
''                mCnn.Execute mSql
''
''                mSql = "Update faBanks Set dtReconStartDate = Null, dtReconEndDate = Null WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
''                mCnn.Execute mSql
''
''                cmdReset.Visible = False
''                Call FormInitialize
''            End If
'
'        Else
'            mResetFlag = 2
'            If MsgBox("Are you sure to reset Reconciliation for the Month " & MonthName(mLastMonth) & " " & mLastYear, vbYesNo) = vbYes Then
'                 'DELETE AND REST DATES IN BANKs TABLE
'                mSql = "Delete From faBankReconcileChild WHERE intReconID = " & mReconID
'                mCnn.Execute mSql
'
'                mSql = "Delete From faBankReconcile WHERE intReconID = " & mReconID
'                mCnn.Execute mSql
'
'                mSql = "Update faBanks Set dtReconEndDate =  '" & DdMmmYy(CDate(mdtLastDate)) & "' WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
'                mCnn.Execute mSql
'
'                UpdateMonthlyStatus (val(lblYear.Tag))
'            Else
'            End If
'
'       End If
'    End If
End Sub

Private Sub cmdStart_Click()
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mArrIn As Variant
    Dim Rec As New ADODB.Recordset
    Dim mReconID As Integer
    Dim mSql As String
    If cmdStart.Caption = "START" Then
        If txtBankStatementAmount.Text = "" Then
            MsgBox "Please Enter Bank statement Amount", vbApplicationModal
            Exit Sub
        End If
    End If
    
    If cmdStart.Caption = "START" Or cmdStart.Caption = "SAVE" Then
        If IsDate(mdtLastDate) Then
            Me.MousePointer = vbHourglass
            Dim mYearID As Integer
            Dim mMonthID As Integer
            mYearID = Year(mdtLastDate)
            mMonthID = Month(mdtLastDate)
            If mMonthID < 4 Then
                mYearID = mYearID - 1
            End If
            If mEditFlag Then
                    mArrIn = Array(txtBankStatementAmount.Tag, _
                    mYearID, _
                    mMonthID, _
                    val(txtHeadCode.Tag), _
                    Trim(txtHeadCode.Text), _
                    val(txtBankBookAmount.Text), _
                    val(txtBankStatementAmount.Text), _
                    0)
            ElseIf cmdStart.Caption = "START" Or cmdStart.Caption = "SAVE" Then
                    mArrIn = Array(-1, _
                    mYearID, _
                    mMonthID, _
                    val(txtHeadCode.Tag), _
                    Trim(txtHeadCode.Text), _
                    val(txtBankBookAmount.Text), _
                    val(txtBankStatementAmount.Text), _
                    0)
            End If
            
            objDb.SetConnection mCnn
            Set Rec = objDb.ExecuteSP("spSaveBankReconcile", mArrIn, , , mCnn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                mReconID = Rec(0).value
                txtBankStatementAmount.Tag = mReconID
                If mEditFlag = True Then
                    mSql = " DELETE FROM faBankReconcileChild FROM faBankReconcileChild " & vbCrLf
                    mSql = mSql + " INNER JOIN faBankreconcile ON faBankReconcile.intReconID = faBankReconcileChild.intReconID " & vbCrLf
                    mSql = mSql + " Where faBankReconcileChild.intReconID = " & mReconID & vbCrLf
                    mSql = mSql + " AND tnyReconStatus = 0 "
                    mCnn.Execute mSql
                End If
                mArrIn = Array(mReconID)
                Call objDb.ExecuteSP("spPortReconTransactions", mArrIn, , , mCnn, adCmdStoredProc)
                mReconStatus = 0
            
            End If
            Rec.Close
            
            '
            'FIRST TIME STARTS TO RECONCILIE SET THE
            ' START_DATA IN BANKs TABLE
            '
            If mFirstTimeReconFlag Then
                If CDate(mdtLastDate) < "01-Apr-2008" Then
                    MsgBox "Error : Start Date", vbInformation
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
               
                mSql = "UPDATE faBanks SET dtReconStartDate = '" & DdMmmYy(CDate(mdtLastDate)) & "'  WHERE intAccountHeadID = " & val(txtHeadCode.Tag)
                mSql = mSql + " AND  isdate(dtReconStartDate) = 0 "
                mCnn.Execute mSql
                
            End If
            'END OF BLOCK
            
            mEditFlag = False
            txtBankStatementAmount.Enabled = False
            cmdStart.Caption = "EDIT"
            
            Me.MousePointer = vbDefault
            
        End If
    ElseIf cmdStart.Caption = "EDIT" Then
        Dim mStr As String
        Dim mBankBalance As Double
        Dim objLdgr As New clsAccounts
        
        
        mStr = "          Do you want to edit the Balance as per Bank Statement?          " & vbCrLf
        mStr = mStr + "[This will Reset all reconciled item of the selected Bank's current month]"

        If MsgBox(mStr, vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            If val(txtHeadCode.Tag) > 0 And IsDate(lblLastDate.Caption) Then
                mBankBalance = objLdgr.GetLedgerBalance(val(txtHeadCode.Tag), CDate(lblLastDate.Caption))
            Else
                mStr = "Didn't able to check Bank Book Balance!" & vbCrLf
                mStr = mStr + " Please try to Edit again by Selecting the Bank and Month properly!" & vbCrLf
                MsgBox mStr, vbInformation
                Exit Sub
            End If
            txtBankBookAmount.ToolTipText = ""
            txtBankBookAmount.BackColor = &HE0E0E0
            txtBankBookAmount.Text = Format(mBankBalance, "0.00")
            txtBankStatementAmount.Enabled = True
            cmdStart.Caption = "SAVE"
            txtBankStatementAmount.SelStart = 0
            txtBankStatementAmount.SelLength = Len(txtBankStatementAmount.Text)
            txtBankStatementAmount.SetFocus
            mEditFlag = True
        End If
    End If
    
End Sub

Private Sub cmdYearDown_Click()
    Call ClearBalanceFields
    If val(lblYear.Tag) < 2009 Then
        lblYear.Tag = 2008
    Else
        lblYear.Tag = val(lblYear.Tag) - 1
    End If
    SetYear (DateSerial(val(lblYear.Tag), 4, 30))
    If val(txtHeadCode.Tag) > 0 Then
        UpdateMonthlyStatus (val(lblYear.Tag))
    End If
End Sub


Private Sub cmdYearUp_Click()
    Call ClearBalanceFields
    If val(lblYear.Tag) < gbFinancialYearID Then
        lblYear.Tag = val(lblYear.Tag) + 1
    Else
        lblYear.Tag = gbFinancialYearID
    End If
    Call SetYear(DateSerial(val(lblYear.Tag), 4, 30))
    
'''''''    Added On 24/Jun/2016 By Anisha
    If val(vsGrid.TextMatrix(vsGrid.Row, 0)) > 0 Then
        Call DisplayBankDetails(val(vsGrid.TextMatrix(vsGrid.Row, 0)))
    End If
''''''    ----------
    'Call SetBalance(CDate(mNextDate))
    If val(txtHeadCode.Tag) > 0 Then
        UpdateMonthlyStatus (val(lblYear.Tag))
    End If
End Sub



Private Sub Form_Load()

    vsGrid.Rows = 20
    vsGrid.Cell(flexcpFontSize, 0) = 12
    vsGrid.RowHeight(0) = 600
    vsGrid.Cell(flexcpFontSize, 0, 0, 0, vsGrid.Cols - 1) = 10
    vsGrid.Cell(flexcpFontBold, 0, 0, 0, vsGrid.Cols - 1) = True
    'vsGrid.Height = vsGrid.Rows * vsGrid.RowHeight(1) + 400
    
    
    vsGridMonth.Rows = 13
    
    vsGridMonth.Cell(flexcpFontSize, 0) = 12
    vsGridMonth.RowHeight(0) = 600
    vsGridMonth.Cell(flexcpFontSize, 0, 0, 0, vsGridMonth.Cols - 1) = 10
    vsGridMonth.Cell(flexcpFontBold, 0, 0, 0, vsGridMonth.Cols - 1) = True
    
    vsGridMonth.TextMatrix(1, 0) = 4
    vsGridMonth.TextMatrix(1, 1) = "APRIL"
    
    vsGridMonth.TextMatrix(2, 0) = 5
    vsGridMonth.TextMatrix(2, 1) = "MAY"
    
    vsGridMonth.TextMatrix(3, 0) = 6
    vsGridMonth.TextMatrix(3, 1) = "JUNE"
    
    vsGridMonth.TextMatrix(4, 0) = 7
    vsGridMonth.TextMatrix(4, 1) = "JULY"
    
    vsGridMonth.TextMatrix(5, 0) = 8
    vsGridMonth.TextMatrix(5, 1) = "AUGUST"
    
    vsGridMonth.TextMatrix(6, 0) = 9
    vsGridMonth.TextMatrix(6, 1) = "SEPTEMBER"
    
    vsGridMonth.TextMatrix(7, 0) = 10
    vsGridMonth.TextMatrix(7, 1) = "OCTOBER"
    
    vsGridMonth.TextMatrix(8, 0) = 11
    vsGridMonth.TextMatrix(8, 1) = "NOVEMBER"
    
    vsGridMonth.TextMatrix(9, 0) = 12
    vsGridMonth.TextMatrix(9, 1) = "DECEMBER"
    
    vsGridMonth.TextMatrix(10, 0) = 1
    vsGridMonth.TextMatrix(10, 1) = "JANUARY"
    
    vsGridMonth.TextMatrix(11, 0) = 2
    vsGridMonth.TextMatrix(11, 1) = "FEBRUARY"
    
    vsGridMonth.TextMatrix(12, 0) = 3
    vsGridMonth.TextMatrix(12, 1) = "MARCH"
    Call FillGird
    SetYear (gbTransactionDate)
End Sub



Private Sub imgNext_Click()
    If val(txtHeadCode.Tag) > 0 Then
    If IsDate(lblLastDate.Caption) Then
    If IsNumeric(txtBankStatementAmount.Text) Then
    If val(txtBankStatementAmount.Tag) > 0 Then ' [faBankReconciliation.intReconID]
        frmReconciliation.ReconciliationStarted = mReconciliationStarted
        frmReconciliation.LastDate = lblLastDate.Caption
        frmReconciliation.BankAccountHeadID = val(txtHeadCode.Tag)
        frmReconciliation.BankBalance = val(txtBankBookAmount.Text)
        frmReconciliation.ReconID = val(txtBankStatementAmount.Tag)
        frmReconciliation.PassBookBalance = val(txtBankStatementAmount.Text)
        frmReconciliation.ReconcileStatus = mReconStatus
        frmReconciliation.Visible = True
        frmReconciliation.ZOrder (0)
    End If
    End If
    End If
    End If
End Sub

Private Sub vsGrid_Click()
    'frmReconMonths.Visible = True
End Sub
Private Sub vsGrid_DblClick()
    If vsGrid.TextMatrix(vsGrid.Row, 0) > 0 Then
        Call FormInitialize
        Call DisplayBankDetails(val(vsGrid.TextMatrix(vsGrid.Row, 0)))
        
         If Not IsDate(mNextDate) Then
            'NEXT DATE MUST BE SET FIRST
            Exit Sub
            '
        End If
        
        Call SetYear(CDate(mNextDate))
        If mFirstTimeReconFlag = False Then
            Call SetBalance(CDate(mNextDate))
        End If
    End If
End Sub

Private Sub DisplayBankDetails(mBankID As Integer)
    Dim objBank As New clsBank
    
    Dim mCnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCount As Integer
    
    lblLastDate.Caption = "#"
    'cmdStart.Enabled = False
    objBank.SetBankInfo (mBankID)
    If objBank.BankID > 0 Then
        txtHeadCode.Tag = objBank.BankAccountHeadID
        txtHeadCode.Text = objBank.BankAccountHeadCode
        txtBankName.Text = objBank.BankName
        
        txtType.Text = objBank.MinorAccountHead
        txtType.Tag = objBank.MinorAccountHeadID
        
        
        ' FINDING NEXT RECONCILIATION DATE :: LAST DATE OF MONTH
        If IsDate(objBank.ReconciliationLastDate) Then
            mNextDate = CDate(objBank.ReconciliationLastDate)
            mNextDate = DateSerial(Year(mNextDate), Month(mNextDate), 1)
            mNextDate = DateAdd("m", 2, mNextDate)
            mNextDate = DateAdd("d", -1, mNextDate)
        ElseIf IsDate(objBank.ReconciliationStartDate) Then
            mNextDate = CDate(objBank.ReconciliationStartDate)
            mFirstTimeReconFlag = True
            cmdReset.Visible = True
        Else
            mNextDate = DateSerial(gbFinancialYearID, 4, 30)
            mFirstTimeReconFlag = True
        End If
        SetYear (mNextDate)
        mdtLastDate = mNextDate
        
        '=========================================='
         Call UpdateMonthlyStatus(val(lblYear.Tag))
        '=========================================='
    Else
        Call FormInitialize
    End If
End Sub

Private Sub FillGird()
    Dim mCnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCount As Integer
    
    objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSql = "Select intBankID,vchBankName,faAccountHeads.vchAccountHeadCode AccHeadCode from faBanks "
    mSql = mSql + " Inner Join faAccountHeads On faBanks.intAccountHeadID=faAccountHeads.intAccountHeadID"
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    vsGrid.Rows = 1
    mRowCount = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intBankID), "", Rec!intBankID)
        vsGrid.TextMatrix(mRowCount, 1) = Rec!AccHeadCode
        vsGrid.TextMatrix(mRowCount, 2) = Rec!vchBankName
        mRowCount = mRowCount + 1
        Rec.MoveNext
    Wend
    Rec.Close
    mCnn.Close
End Sub

Private Sub vsGridMonth_Click()
    vsGridMonth.HighLight = flexHighlightAlways
End Sub

Private Sub vsGridMonth_DblClick()
    
    Dim mDt As Date
    If val(txtHeadCode.Tag) > 0 Then  ' NOTE: IF BANK IS SELECTED THEN
        Dim mYearID As Integer
        Dim mMonthID As Integer
        mYearID = val(lblYear.Tag)
        mMonthID = val(vsGridMonth.TextMatrix(vsGridMonth.Row, 0))
        If mMonthID < 4 Then
            mYearID = mYearID + 1
        End If
            
        mDt = DateSerial(mYearID, mMonthID, 1)
        mDt = FindLastDateOfTheMonth(mDt)
        SetBalance (mDt)
        UpdateMonthlyStatus (val(lblYear.Tag))
    
    End If
End Sub
