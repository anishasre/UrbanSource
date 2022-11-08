VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmBankScrollOpening 
   BackColor       =   &H00EDF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "             Bank Scroll/Passbook Opening Entries"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankScrollOpening.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtActualDiff 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   10095
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1590
   End
   Begin VB.TextBox txtPassBookBalance 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1785
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6615
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1785
   End
   Begin VB.TextBox txtDifference 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   10095
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5775
      Visible         =   0   'False
      Width           =   1590
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   13860
      Top             =   7110
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtLedgerBookBalance 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2220
   End
   Begin VB.CommandButton cmdSearchBank 
      Caption         =   "..."
      Height          =   285
      Left            =   8610
      TabIndex        =   3
      Top             =   420
      Width           =   330
   End
   Begin VB.TextBox txtBank 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   5910
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11850
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDF7F7&
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   1050
      Width           =   11610
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EDF7F7&
      Height          =   195
      Left            =   15
      TabIndex        =   8
      Top             =   5490
      Width           =   11610
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EDF7F7&
      Height          =   690
      Left            =   45
      TabIndex        =   17
      Top             =   6165
      Width           =   11670
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00EDF7F7&
         Caption         =   "&Save"
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
         Left            =   5310
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00EDF7F7&
         Caption         =   "C&lose"
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
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00EDF7F7&
         Caption         =   "&Clear"
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
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   225
         Width           =   1185
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3750
      Left            =   45
      TabIndex        =   7
      Top             =   1740
      Width           =   11655
      _cx             =   20558
      _cy             =   6615
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15595511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15595511
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBankScrollOpening.frx":000C
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
      Editable        =   2
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
   Begin VB.Label Label6 
      BackColor       =   &H00EDF7F7&
      Height          =   330
      Left            =   2055
      TabIndex        =   21
      Top             =   750
      Width           =   6885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount to be Adjusted"
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   8010
      TabIndex        =   15
      Top             =   1395
      Width           =   2010
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Differnce "
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9270
      TabIndex        =   13
      Top             =   5820
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculated Pass Book Balance"
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Top             =   5775
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passbook Balance"
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   5820
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Opening Balance"
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   1320
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   240
      Left            =   2130
      TabIndex        =   1
      Top             =   465
      Width           =   465
   End
End
Attribute VB_Name = "frmBankScrollOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mHeadCode            As Variant
    Dim mAccHead             As Variant
    Dim mintID               As Variant
    Dim mMinDate             As Date
    
    Private Sub cmdClear_Click()
        Call formInitialise
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    Private Sub cmdSave_Click()
        Dim mLoop           As Long
        Dim objDB           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim arrInput        As Variant
        Dim mSql            As String
        Dim mDrCrFlag       As Integer
        Dim mVoucherNo      As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        '       Validation  '
        If Trim(txtBank.Text) = "" Then
            MsgBox "Please Select Bank Account", vbApplicationModal
            txtBank.SetFocus
            Exit Sub
        End If
        If val(vsGrid.TextMatrix(1, 0)) = 0 And vsGrid.TextMatrix(1, 2) = "" And val(vsGrid.TextMatrix(1, 3)) = 0 Then
            Exit Sub
        End If
'        If vsGrid.TextMatrix(vsGrid.Row, 3) = 0 Or vsGrid.TextMatrix(vsGrid.Row, 3) < 0 Then
'            MsgBox "PLease Enter the Amount", vbInformation
'            Exit Sub
'        End If
        '---------------------'
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 0) = "" Then
                Exit For
            End If
            If vsGrid.TextMatrix(mLoop, 2) = "" Then
                Exit For
            End If
            If vsGrid.TextMatrix(mLoop, 3) = "" Then
                Exit For
            End If
            
            If val(vsGrid.TextMatrix(mLoop, 13)) <> 1 Then      '' Checking If Not a Difference amount
                 If vsGrid.Cell(flexcpValue, mLoop, 2) = 1 Then      '' Cheque Issued
                     mDrCrFlag = 0 '' Credit
                 ElseIf vsGrid.Cell(flexcpValue, mLoop, 2) = 2 Then  '' Cheque Deposited
                     mDrCrFlag = 1 '' Debit
                 ElseIf vsGrid.Cell(flexcpValue, mLoop, 2) = 3 Then  '' Directly Debited
                     mDrCrFlag = 1 '' Debit
                 Else                                                '' Directly Credited
                     mDrCrFlag = 0
                 End If
                 
                 If vsGrid.TextMatrix(mLoop, 1) <> "" Then           '' Voucher Number
                     mVoucherNo = vsGrid.TextMatrix(mLoop, 2)
                     mVoucherNo = Token(mVoucherNo, "/")
                 End If
                
                 mintID = IIf(vsGrid.TextMatrix(mLoop, 8) = "", -1, vsGrid.TextMatrix(mLoop, 8))
                 If vsGrid.Cell(flexcpValue, mLoop, 2) = 1 Or vsGrid.Cell(flexcpValue, mLoop, 2) = 2 Then
                     arrInput = Array(mintID, _
                                 Null, _
                                 vsGrid.TextMatrix(mLoop, 9), _
                                 val(mVoucherNo), _
                                 val(vsGrid.TextMatrix(mLoop, 4)), _
                                 vsGrid.TextMatrix(mLoop, 5), _
                                 vsGrid.Cell(flexcpValue, mLoop, 2) * 10, _
                                 vsGrid.TextMatrix(mLoop, 7), _
                                 vsGrid.TextMatrix(mLoop, 11), _
                                 Null, _
                                 vsGrid.TextMatrix(mLoop, 0), _
                                 mDrCrFlag, _
                                 vsGrid.TextMatrix(mLoop, 3), _
                                 Null, _
                                 vsGrid.TextMatrix(mLoop, 6), _
                                 val(txtBank.Tag), _
                                 mHeadCode, _
                                 Null, _
                                 vsGrid.TextMatrix(mLoop, 6), _
                                 vsGrid.TextMatrix(mLoop, 1), _
                                 Trim(txtBank.Text), _
                                 vsGrid.TextMatrix(mLoop, 12))
                     objDB.ExecuteSP "spSaveOpeningVouchers", arrInput, , , mCnn
                 Else
                     arrInput = Array(IIf(mintID < 0, Null, mintID), _
                                 val(txtBank.Tag), _
                                 mHeadCode, _
                                 vsGrid.TextMatrix(mLoop, 0), _
                                 vsGrid.TextMatrix(mLoop, 6), _
                                 vsGrid.TextMatrix(mLoop, 4), _
                                 vsGrid.TextMatrix(mLoop, 5), _
                                 IIf(mDrCrFlag = 1, vsGrid.TextMatrix(mLoop, 3), 0), _
                                 IIf(mDrCrFlag = 0, vsGrid.TextMatrix(mLoop, 3), 0), _
                                 1)
                     objDB.ExecuteSP "spSaveBankReconsilation", arrInput, , , mCnn
                 End If
            End If
        Next mLoop
        
        '--------Diference Amount Saving----------'
        mCnn.Execute "Delete From faBankReconciliationEntries Where tnyType = 1 And tnyOpening = 1 And intBankAccountHeadID = " & val(txtBank.Tag)
        If val(txtActualDiff.Text) <> 0 Then
            If val(txtPassBookBalance.Text) - val(txtTotal.Text) < 0 Then
                mDrCrFlag = 1
            Else
                mDrCrFlag = 0
            End If
            mintID = -1
            arrInput = Array(IIf(mintID < 0, Null, mintID), _
                            val(txtBank.Tag), _
                            mHeadCode, _
                            mMinDate, _
                            "Difference Amount", _
                            Null, _
                            Null, _
                            IIf(mDrCrFlag = 1, val(txtActualDiff.Text), 0), _
                            IIf(mDrCrFlag = 0, val(txtActualDiff.Text), 0), _
                            1, _
                            1) '' Last Parameter for Difference Amount
            objDB.ExecuteSP "spSaveBankReconsilation", arrInput, , , mCnn
        End If
        '------------------------------------------'
        MsgBox "Saved Successfully!", vbInformation, "Saankhya"
        cmdSave.Enabled = False
        Call FillGrid
    End Sub
    '-----Sinoj------'
    Private Sub cmdSearchBank_Click()
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where intGroupID = 2 And tinHiddenFlag = 0"
        frmSearchAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            txtBank.Tag = gbSearchID
            txtBank.Text = Trim(gbSearchStr)
            mHeadCode = Token(gbSearchStr, " ")
            mAccHead = Split(Trim(txtBank.Text), " ")
            Call getBankLedgerBalance
            Call getBankPassBookBalance
            Call FillGrid
            txtTotal.Text = Format(getTotalAmounts(), "0.00")
            txtActualDiff.Text = Format(Abs(val(txtPassBookBalance.Text) - val(txtTotal.Text)), "0.00")
'            If val(txtDifference.Text) = val(txtActualDiff.Text) Then
'                txtActualDiff.Text = ""
'            End If
            cmdSearchBank.Enabled = False
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub Form_Load()
        Dim mComboStr As String
        Dim mArrOut, Cnt As Integer
        Dim mSql        As String
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mLoop       As Integer
        winXPC.InitIDESubClassing
        
        mMinDate = gbStartingDate - 1
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select min(dtTransactionDate) MiDate From faTransactions"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mMinDate = Rec!mIDate
        End If
        Rec.Close
        
        Call formInitialise
        mComboStr = ""
        mComboStr = mComboStr & "#1;Cheque Issued(+)|"
        mComboStr = mComboStr & "#2;Cheque Deposited(-)|"
        mComboStr = mComboStr & "#3;Debited by Bank(-)|"
        mComboStr = mComboStr & "#4;Credited by Bank(+)"
        
        vsGrid.ColComboList(2) = mComboStr
        Label6.Visible = True
        Label6.Caption = " Please Enter the Voucher date before  " & mMinDate
    End Sub
    Private Sub formInitialise()
        txtBank.Text = ""
        txtBank.Tag = -1
        txtLedgerBookBalance.Text = ""
        txtPassBookBalance.Text = ""
        txtDifference.Text = ""
        txtTotal.Text = ""
        txtActualDiff.Text = ""
        vsGrid.Rows = 1
        vsGrid.Rows = 2
        cmdSave.Enabled = True
        cmdSearchBank.Enabled = True
    End Sub
    '-----Sinoj------'
    Private Sub getBankLedgerBalance()
        Dim mSql As String
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        If txtBank.Tag <> -1 Then
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
''            mSQL = "Select faTransactionChild.fltAmount*((tinDebitOrCreditFlag*+2)-1)fltAmount" & vbNewLine & _
''                    "From faTransactions" & vbNewLine & _
''                    "Inner Join  faTransactionChild  On faTransactionChild.intTransactionID = faTransactions.intTransactionID" & vbNewLine & _
''                    "Where dtTransactionDate < Convert(smallDateTime,Convert(varchar(11),getDate())) And isNull(faTransactions.tnyStatus,0) <> 4 And intAccountHeadID = " & txtBank.Tag
            
            mSql = "Select faTransactionChild.fltAmount*((tinDebitOrCreditFlag*+2)-1)fltAmount" & vbNewLine & _
                    "From faTransactions" & vbNewLine & _
                    "Inner Join  faTransactionChild  On faTransactionChild.intTransactionID = faTransactions.intTransactionID" & vbNewLine & _
                    "Where intTransactiontypeID=3000 And  dtTransactionDate < Convert(smallDateTime,Convert(varchar(11),getDate())) And isNull(faTransactions.tnyStatus,0) <> 4 And intAccountHeadID = " & txtBank.Tag
            Rec.Open mSql, mCnn
            txtLedgerBookBalance.Text = ""
            If Not (Rec.EOF And Rec.BOF) Then
                txtLedgerBookBalance.Text = Format(IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount), "0.00")
            End If
            Rec.Close
            mCnn.Close
        End If
    End Sub
    '-----Sinoj-----'
    Private Sub getBankPassBookBalance()
        Dim mSql As String
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mPassBookBalance As Variant
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mPassBookBalance = 0
        txtPassBookBalance.Text = ""
''        mSQL = "Select  isNull(Sum(fltCrAmount),0)-isNull(Sum( fltDrAmount),0) ScrollBalance From faBankReconciliationEntries" & vbNewLine & _
''                "Where tnyOpening = 0 AND intBankAccountHeadID = " & txtBank.Tag & " And dtBankEntryDate <=  Convert(smallDateTime,Convert(smallDateTime,getDate()))"
''
''        mSql = "Select  isNull(Sum(fltCrAmount),0)-isNull(Sum( fltDrAmount),0) ScrollBalance From faBankReconciliationEntries" & vbNewLine & _
''              "Where tnyOpening = 1 AND intBankAccountHeadID = " & txtBank.Tag & " And dtBankEntryDate <=  Convert(smallDateTime,Convert(smallDateTime,getDate()))"
''        Rec.Open mSql, mCnn
''        If Not (Rec.EOF And Rec.BOF) Then
''            mPassBookBalance = Rec!ScrollBalance
''        End If
''        Rec.Close
''        '-- AMOUNT FROM FAoPENING VOUCHERS  Added By Anisha For Checking
''        mSql = "Select Sum(case when tinDebitOrCreditFlag=1 then fltAmount When tinDebitOrCreditFlag=0 then fltAmount*-1 End) as OpeningAmount" & vbNewLine & _
''                " From faOpeningVouchers " & vbNewLine & _
''                "Where intAccountHeadID = " & txtBank.Tag & ""
''        Rec.Open mSql, mCnn
''        If Not (Rec.EOF And Rec.BOF) Then
''            mPassBookBalance = Rec!OpeningAmount
''        End If
''        Rec.Close
        
        mSql = "SELECT fltOpening*((tinDebitOrCreditFlag*+2)-1) fltAmount FROM faBanks Where intAccountHeadID = " & txtBank.Tag
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mPassBookBalance = mPassBookBalance + IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            txtPassBookBalance.Text = Format(mPassBookBalance, "0.00")
        End If
        Rec.Close
        mCnn.Close
    End Sub
    '-----Sinoj------'
    Private Function getTotalAmounts() As Double
        Dim mTotal As Double
        Dim mLoop As Integer
        mTotal = val(txtLedgerBookBalance.Text)
        For mLoop = 1 To vsGrid.Rows - 1
            With vsGrid
                If (.Cell(flexcpValue, mLoop, 2)) = 1 Or (.Cell(flexcpValue, mLoop, 2)) = 4 Then ' Cheque Issued or Directly Debited
                    mTotal = mTotal + val(.TextMatrix(mLoop, 3))
                Else
                    mTotal = mTotal - val(.TextMatrix(mLoop, 3))
                End If
            End With
        Next mLoop
        getTotalAmounts = mTotal
    End Function
    Private Sub FillGrid()
        Dim mSql        As String
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mLoop       As Integer
        
        mMinDate = gbStartingDate - 1
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select min(dtTransactionDate) MiDate From faTransactions"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mMinDate = Rec!mIDate
        End If
        Rec.Close
        mSql = "Select  dtBankEntryDate dtDate,Null intVoucherNo,Case When isNull(fltDrAmount,0) <> 0 Then 3 Else 4 End Type,Case When isNull(fltDrAmount,0) <> 0 Then isNull(fltDrAmount,0) Else isNull(fltCrAmount,0) End fltAmount," & vbNewLine & _
                "vchchequeNo vchInstrumentNo,dtchequeDate dtInstrumentDate,vchParticulars vchRemarks,tnyReconciled,intreconciliationID intID,intVoucherID,tnytype, 0 tbl,numTockenID,dtReconcileDate" & vbNewLine & _
                "From faBankReconciliationEntries" & vbNewLine & _
                "Where tnyOpening = 1 And dtBankEntryDate <= '" & Format(mMinDate, "dd/MMM/YYYY") & "' And intBankAccountHeadID = " & txtBank.Tag & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select dtDate,vchVNo intVoucherNo,tinDebitOrCreditFlag+1 Type,fltAmount,vchInstrumentNo,dtInstrumentDate,vchRemarks,tnyReconciled,intID,intVoucherID,Null tnyType,1 tbl,numTockenID,dtReconcileDate" & vbNewLine & _
                "From faOpeningVouchers Where dtDate <= '" & Format(mMinDate, "dd/MMM/YYYY") & "' And intAccountHeadID = " & txtBank.Tag
        vsGrid.Rows = 1
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Rec.EOF = False
                With vsGrid
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Format(IIf(IsNull(Rec!dtDate), "", Rec!dtDate), "dd-MMM-yyyy")
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    .Cell(flexcpText, .Rows - 1, 2) = IIf(IsNull(Rec!Type), "", Rec!Type)
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    .TextMatrix(.Rows - 1, 5) = Format(IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate), "dd-MMM-yyyy")
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    .TextMatrix(.Rows - 1, 7) = IIf(IsNull(Rec!tnyReconciled), "", Rec!tnyReconciled)
                    .TextMatrix(.Rows - 1, 8) = IIf(IsNull(Rec!intID), "", Rec!intID)
                    .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    .TextMatrix(.Rows - 1, 10) = IIf(IsNull(Rec!tbl), "", Rec!tbl)
                    .TextMatrix(.Rows - 1, 11) = IIf(IsNull(Rec!numTockenID), "", Rec!numTockenID)
                    .TextMatrix(.Rows - 1, 12) = IIf(IsNull(Rec!dtReconcileDate), "", Rec!dtReconcileDate)
                    
                    If IsNull(Rec!tnyType) = False Then    '' For Difference Amount With Pass Book Balance and Ledger Book Balance
                        txtDifference.Text = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
                        .TextMatrix(.Rows - 1, 13) = IIf(IsNull(Rec!tnyType), "", Rec!tnyType)
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, 13) = vbBlue
                    End If
                End With
                Rec.MoveNext
            Wend
        End If
        vsGrid.Rows = vsGrid.Rows + 1
        Rec.Close
    End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim mDiffRow As Integer
        
        Label6.Caption = ""
        Label6.Visible = False
        If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
           vsGrid.TextMatrix(vsGrid.Row, 0) = CheckDateInMMM(vsGrid.TextMatrix(vsGrid.Row, 0))
           If vsGrid.TextMatrix(vsGrid.Row, 5) <> "" Then
               vsGrid.TextMatrix(vsGrid.Row, 5) = CheckDateInMMM(vsGrid.TextMatrix(vsGrid.Row, 5))
           End If
        End If
        '-----Added on 28/3/2011 By Poornima
        If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
           If CDate(vsGrid.TextMatrix(vsGrid.Row, 0)) >= CDate(mMinDate + 1) Or CDate(vsGrid.TextMatrix(vsGrid.Row, 0)) < CDate("01/Jan/1900") Then
'              MsgBox "Please Enter a Valid Date", vbInformation
                MsgBox " Please Enter a Date before  " & mMinDate, vbInformation
                vsGrid.TextMatrix(vsGrid.Row, 0) = ""
                Label6.Visible = True
                Label6.Caption = " Please Enter the date before  " & mMinDate
                                  
           End If
        End If
        
        If vsGrid.TextMatrix(vsGrid.Row, 5) <> "" Then
           If CDate(vsGrid.TextMatrix(vsGrid.Row, 5)) > CDate(gbEndingDate) Or CDate(vsGrid.TextMatrix(vsGrid.Row, 5)) < CDate("01/Jan/1900") Then
              MsgBox "Please Enter a  Date within this Financial Year", vbInformation
              vsGrid.TextMatrix(vsGrid.Row, 5) = ""
           End If
        End If
        If Col = 3 Then
            mDiffRow = vsGrid.FindRow(1, 1, 13)
            If mDiffRow > 0 Then
                vsGrid.TextMatrix(mDiffRow, 3) = 0
            End If
        End If
        '''----------------Added by poornima on 28/07/2011----------------------''
        vsGrid.TextMatrix(vsGrid.Row, 3) = Format(val(vsGrid.TextMatrix(vsGrid.Row, 3)), "0.00")
        If Len(vsGrid.TextMatrix(vsGrid.Row, 3)) > 15 Then
            MsgBox "Please check the Amount", vbInformation
            vsGrid.TextMatrix(vsGrid.Row, 3) = ""
        End If
        If Len(vsGrid.TextMatrix(vsGrid.Row, 4)) > 15 Then
            MsgBox "Please check the Number", vbInformation
            vsGrid.TextMatrix(vsGrid.Row, 4) = ""
        End If
        ''------------------------------------------------------------------------'
        '----Calculations and Display------
        txtTotal.Text = Format(getTotalAmounts(), "0.00")
        txtActualDiff.Text = Format(Abs(val(txtPassBookBalance.Text) - val(txtTotal.Text)), "0.00")
'        If val(txtDifference.Text) = val(txtActualDiff.Text) Then
'            txtActualDiff.Text = ""
'        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If txtBank.Tag < 1 Then
            MsgBox "Please Select the Bank", vbInformation
            Exit Sub
        End If
        If vsGrid.TextMatrix(vsGrid.Row, 7) <> "" Then
            MsgBox "Already Reconcilled and cannot be edited", vbApplicationModal
            Exit Sub
        End If
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Col = 0 Or Col = 5 Then
            Call KeyPressNumber(KeyAscii, "/-")
        ElseIf Col = 1 Then
            Call KeyPressNumber(KeyAscii, "/") ', True)
        ElseIf Col = 3 Then
            Call KeyPressNumber(KeyAscii, ".")
        ElseIf Col = 4 Then
            Call KeyPressNumber(KeyAscii, "/-") ', True)
        End If
        
        If vsGrid.TextMatrix(Row, 13) = "1" Then
            If Col = 3 Then
                Call KeyPressNumber(KeyAscii) ', , , "123456789")
            End If
        End If
        '------------------------------------------------------------------------
        If vsGrid.Col = 6 Then
            vsGrid.EditMaxLength = 15
        End If
        '------------------------------------------------------------------------
        '----------New Row Addition----------'
        If vsGrid.Rows - 1 = Row And Col = 6 And KeyAscii = 13 Then
           If Trim(vsGrid.TextMatrix(Row, 0)) <> "" And Trim(vsGrid.TextMatrix(Row, 2)) <> "" And val(Trim(vsGrid.TextMatrix(Row, 3))) <> 0 Then
              vsGrid.Rows = vsGrid.Rows + 1
              vsGrid.Row = vsGrid.Rows - 1
              vsGrid.Col = 0
           End If
        End If
    End Sub
   
