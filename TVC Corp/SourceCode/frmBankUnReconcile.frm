VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBankUnReconcile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UnReconcile"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13320
   Icon            =   "frmBankUnReconcile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7252
      TabIndex        =   16
      Top             =   6885
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   5047
      TabIndex        =   15
      Top             =   6885
      Width           =   1095
   End
   Begin VB.CommandButton cmdUnReconcile 
      Caption         =   "UnReconcile"
      Height          =   375
      Left            =   6127
      TabIndex        =   14
      Top             =   6885
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   135
      TabIndex        =   1
      Top             =   450
      Width           =   13155
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   12285
         TabIndex        =   18
         Top             =   270
         Width           =   690
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmBankUnReconcile.frx":1CCA
         Left            =   10530
         List            =   "frmBankUnReconcile.frx":1CF5
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   315
         Width           =   1680
      End
      Begin VB.CommandButton cmdYearDown 
         BackColor       =   &H8000000B&
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   315
         Width           =   285
      End
      Begin VB.CommandButton cmdYearUp 
         BackColor       =   &H8000000B&
         Caption         =   "Ù"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9495
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   315
         Width           =   285
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         TabIndex        =   6
         Top             =   315
         Width           =   825
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "..."
         Height          =   330
         Left            =   7155
         TabIndex        =   3
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txtBank 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   5820
      End
      Begin VB.Label lblLast 
         Caption         =   "Last Reconcile Month"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   19
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9990
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7785
         TabIndex        =   10
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5280
      Left            =   90
      TabIndex        =   0
      Top             =   1530
      Width           =   13200
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   4695
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   13110
         _cx             =   23125
         _cy             =   8281
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   14609886
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483624
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
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
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmBankUnReconcile.frx":1D35
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   4
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
      Begin VB.Label Label3 
         Caption         =   "Reconciled Bank Statement"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   180
         Width           =   2400
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UNRECONCILE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2025
      TabIndex        =   13
      Top             =   45
      Width           =   9240
   End
   Begin VB.Label Label1 
      BackColor       =   &H0098B7A3&
      Height          =   420
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   13200
   End
End
Attribute VB_Name = "frmBankUnReconcile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mStartYear As Integer
    Dim mEndYear  As Integer
    Dim mLastReconmonth  As Integer
    Dim mLastReconYear  As Integer
    Dim mLastReconDate  As Date
    Dim mSelectedMonth  As Date
    Dim mLastReconciledMonth As Date
    Private Sub cmdBank_Click()
        Dim objAcc      As New clsAccounts
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where faAccountHeads.intGroupID =2 And tinHiddenFlag = 0"
        frmSearchAccountHeads.Show vbModal
        txtBank.Tag = gbSearchID
        txtBank.SetFocus
        objAcc.SetAccountID (txtBank.Tag)
        txtBank.Text = objAcc.AccountCode + " " + objAcc.AccountHead
        
    End Sub
    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdSearch_Click()
        Dim mSql        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mSDate       As Date
        Dim mEDate       As Date
        Dim mCnt        As Integer
        Dim mDateFilter As String
        Dim mLast       As String
        
        If txtBank.Text = "" Then
            MsgBox "Please Select Bank", vbInformation
            Exit Sub
        End If
        If txtYear.Text = "" Then
            MsgBox "Please Select Year", vbInformation
            Exit Sub
        End If
        If cmbMonth.ListIndex = -1 Then
            MsgBox "Please Select Month", vbInformation
            Exit Sub
        End If
           
        mDateFilter = "1/" & cmbMonth.ItemData(cmbMonth.ListIndex) & "/" & txtYear.Text
        mSDate = CDate(mDateFilter)
        mEDate = DateAdd("m", 1, mSDate)
        mEDate = DateAdd("d", -1, mEDate)
        
        mLastReconmonth = cmbMonth.ItemData(cmbMonth.ListIndex)
        mLastReconYear = txtYear.Text
        mLastReconDate = "1/" & mLastReconmonth & "/" & mLastReconYear
        mLastReconDate = DateAdd("m", 1, mLastReconDate)
        mLastReconDate = CDate(DateAdd("d", -1, mLastReconDate))
        mSelectedMonth = mEDate
        mSql = "Select Max(dtBankEntryDate) LDate"
        mSql = mSql + " From faBankReconciliationEntries Where intBankAccountHeadID=" & txtBank.Tag & " And tnyReconciled is not Null"
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            Rec.Open mSql, mCnn
            If Not (Rec.BOF And Rec.EOF) Then
                If Not IsNull(Rec!LDate) Then
                    mLast = "1/" & Month(Rec!LDate) & "/" & Year(Rec!LDate)
                    mLastReconciledMonth = DateAdd("m", 1, mLast)
                    mLastReconciledMonth = CDate(DateAdd("d", -1, mLastReconciledMonth))
                    mLast = IIf(IsNull(Rec!LDate), "", MonthName((Month(Rec!LDate))) & " - " & Year(Rec!LDate))
                    lblLast.Caption = "Last Reconciled Month :- " & mLast '& IIf(IsNull(Rec!LDate), "", MonthName((Month(Rec!LDate))) & " - " & Year(Rec!LDate))
                    cmdUnReconcile.Enabled = True
                Else
                    'mLastReconmonth = Null
                    lblLast.Caption = "Reconcilation not Started "
                    cmdUnReconcile.Enabled = False
                End If
            End If
            Rec.Close
        End If
        mCnn.Close
        
        If txtBank.Text <> "" Then
            mSql = "Select dtBankEntrydate,vchParticulars,vchChequeNo,dtChequeDate,fltDrAmount,fltcrAmount,intVoucherNo,intReconciliationID " & vbNewLine
            mSql = mSql + " From faBankReconciliationEntries Where intBankAccountHeadID=" & txtBank.Tag & " And tnyReconciled is not Null" & vbNewLine
            mSql = mSql + " And dtBankEntrydate Between '" & Format(mSDate, "dd-mmm-yyyy") & "' And '" & Format(mEDate, "dd-mmm-yyyy") & "'" & vbNewLine
            If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
               ' Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                Rec.CursorLocation = adUseClient
                Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
                vsGrid.Rows = 1
                If Not (Rec.BOF And Rec.EOF) Then
                    vsGrid.Rows = Rec.RecordCount + 1
                    vsGrid.Col = 1
                    vsGrid.Row = 1
                    vsGrid.ColSel = 8
                    vsGrid.RowSel = vsGrid.Rows - 1
                    mSql = Rec.GetString(, , vbTab, Chr(13))
                    vsGrid.Clip = mSql
                End If
                Rec.Close
            End If
         End If
    End Sub

    Private Sub cmdNew_Click()
        txtBank.Tag = -1
        txtBank.Text = ""
        vsGrid.Clear 1, 0
        
    End Sub
    Private Sub cmdUnReconcile_Click()
        Dim mCnt        As Integer
        Dim mSql        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mRec         As New ADODB.Recordset
        Dim mVrID       As Double
        Dim mTrID       As Double
        Dim mMode       As Integer '' To identyfy FaVoucher or openingVoucher table
        If mLastReconciledMonth = mSelectedMonth Then
        
          If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
    
              For mCnt = 1 To vsGrid.Rows - 1
                  If vsGrid.TextMatrix(mCnt, 7) <> "" Then
                      Rec.Open "Select * From faTransactions Where intVoucherID=(Select intVoucherId From faVouchers Where intVoucherNo=" & vsGrid.TextMatrix(mCnt, 7) & ")", mCnn
                      If Not (Rec.EOF And Rec.BOF) Then
                          mMode = 0
                          mTrID = Rec!intTransactionID
                          mVrID = Rec!intVoucherID
                      Else
                          mRec.Open "Select * From faOpeningVouchers Where numTockenID=" & vsGrid.TextMatrix(mCnt, 8) & "And intVoucherNo=" & vsGrid.TextMatrix(mCnt, 7) & ")", mCnn
                          mMode = 1
                          mVrID = mRec!intID
                      End If
                      Rec.Close
                  End If
                  mSql = "Update faBankReconciliationEntries set tnyReconciled=null,intVoucherNo=Null,dtReconcileDate=Null "
                  mSql = mSql + " Where intReconciliationID=" & val(vsGrid.TextMatrix(mCnt, 8))
                  mCnn.Execute mSql
                  If mMode = 0 Then
                      mSql = "Update faVouchers set tnysync=Null,numTockenID=null,tnyReconciled=Null"
                      mSql = mSql + " Where intVoucherNo=" & mVrID
                      mCnn.Execute mSql
                      
                      mSql = "Update faTransactionChild set tnysync=Null,numTockenID=null,dtReconcileDate=Null"
                      mSql = mSql + " Where intAccountHeadID =" & val(txtBank.Tag) & " And intTransactionID= " & mTrID
                      mCnn.Execute mSql
                  Else
                      mSql = "Update faOpeningVouchers set tnyReconciled=null,numTockenID=Null,dtReconcileDate=Null"
                      mSql = mSql + " Where intAccountHeadID =" & val(txtBank.Tag) & " And intID= " & mVrID
                      mCnn.Execute mSql
                  End If
              Next
              
          End If
        Else
            MsgBox "Please UnReconcile From Last Reconciled Month", vbApplicationModal
            Exit Sub
        End If
        Call cmdSearch_Click
    End Sub

    Private Sub cmdYearDown_Click()
        If txtYear = mStartYear Then
            cmdYearDown.Enabled = False
        Else
            cmdYearUp.Enabled = True
            txtYear.Text = val(txtYear.Text) - 1
        End If
    End Sub

    Private Sub cmdYearUp_Click()
        If txtYear = mEndYear + 1 Then
            cmdYearUp.Enabled = False
        Else
            cmdYearDown.Enabled = True
            txtYear.Text = val(txtYear.Text) + 1
        End If
    End Sub

    Private Sub Form_Load()
        Call FinancialYear
        txtYear.Text = gbFinancialYearID
        
    End Sub
    Private Sub FinancialYear()
        Dim mSql    As String
        Dim objDb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
    
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select min(intFinancialYear) mSYear,max(intFinancialYear) mLYear From faFinancialYear"
            Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                mStartYear = Rec!mSYear
                mEndYear = Rec!mLYear
            End If
        End If
    End Sub

