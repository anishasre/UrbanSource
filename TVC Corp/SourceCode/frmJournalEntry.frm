VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmJournalEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  J o u r n a l "
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmJournalEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   6210
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5325
      TabIndex        =   19
      Top             =   6210
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   6570
      TabIndex        =   21
      Top             =   6210
      Width           =   1215
   End
   Begin VB.ListBox lstMasters 
      Height          =   2205
      Left            =   4635
      TabIndex        =   41
      Top             =   630
      Width           =   2700
   End
   Begin VB.Frame Frame1 
      Height          =   6075
      Left            =   0
      TabIndex        =   37
      Top             =   60
      Width           =   11790
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   150
         Width           =   1575
      End
      Begin VB.Frame fraAdjustments 
         Enabled         =   0   'False
         Height          =   690
         Left            =   585
         TabIndex        =   47
         Top             =   3105
         Width           =   10725
         Begin VB.CheckBox chkRP 
            Caption         =   "Affecting Receipt Payment Statement"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   180
            TabIndex        =   50
            Top             =   180
            Width           =   3615
         End
         Begin VB.CommandButton cmdSearchTransactions 
            Caption         =   "..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8550
            TabIndex        =   49
            Top             =   270
            Width           =   330
         End
         Begin VB.TextBox txtRPLink 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6750
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   270
            Width           =   1770
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Receipt No For Adjustments "
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4275
            TabIndex        =   51
            Top             =   315
            Width           =   2430
         End
      End
      Begin VB.CommandButton cmdSearchVoucherNo 
         Height          =   285
         Left            =   11400
         Picture         =   "frmJournalEntry.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   345
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   3
         Top             =   150
         Width           =   1665
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7110
         TabIndex        =   2
         Top             =   150
         Width           =   1485
      End
      Begin VB.ComboBox cmbTransactionType 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   2505
      End
      Begin VB.Frame fraBudget 
         Height          =   1110
         Left            =   585
         TabIndex        =   22
         Top             =   390
         Width           =   10710
         Begin VB.CommandButton cmdFund 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5055
            TabIndex        =   6
            Top             =   465
            Width           =   315
         End
         Begin VB.CommandButton cmdField 
            Caption         =   "..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9375
            TabIndex        =   13
            Top             =   780
            Width           =   315
         End
         Begin VB.CommandButton cmdFunction 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5055
            TabIndex        =   9
            Top             =   780
            Width           =   315
         End
         Begin VB.CommandButton cmdFunctionary 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9375
            TabIndex        =   11
            Top             =   465
            Width           =   315
         End
         Begin VB.TextBox txtFund 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   465
            Width           =   2550
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6795
            TabIndex        =   12
            Top             =   780
            Width           =   2550
         End
         Begin VB.TextBox txtFunction 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   780
            Width           =   2550
         End
         Begin VB.TextBox txtFunctionary 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6795
            TabIndex        =   10
            Top             =   465
            Width           =   2550
         End
         Begin VB.TextBox txtBudgetCentreCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   135
            TabIndex        =   25
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fund"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1995
            TabIndex        =   29
            Top             =   495
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Field"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6330
            TabIndex        =   28
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Function"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1665
            TabIndex        =   27
            Top             =   795
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Functionary"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5655
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bedget Centre Code"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   165
            TabIndex        =   24
            Top             =   450
            Width           =   1665
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            Caption         =   "  Budget Centre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   255
            Left            =   30
            TabIndex        =   23
            Top             =   120
            Width           =   10635
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1740
         Left            =   585
         TabIndex        =   38
         Top             =   1395
         Width           =   10710
         Begin VB.TextBox txtClaiment 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6465
            MultiLine       =   -1  'True
            TabIndex        =   45
            Text            =   "frmJournalEntry.frx":1DC4
            Top             =   1230
            Width           =   3750
         End
         Begin VB.CommandButton cmdSubLedger 
            BackColor       =   &H00C8E7E7&
            Caption         =   "..."
            Height          =   300
            Left            =   10245
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   840
            Width           =   285
         End
         Begin VB.TextBox txtSubsidiaryLedger 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6465
            TabIndex        =   43
            Top             =   840
            Width           =   3750
         End
         Begin VB.CommandButton cmdSearchAccountHead 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10215
            TabIndex        =   14
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox txtAccountHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4740
            TabIndex        =   40
            Top             =   525
            Width           =   5460
         End
         Begin VB.TextBox txtNarration 
            Appearance      =   0  'Flat
            Height          =   540
            Left            =   2340
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   1140
            Width           =   3210
         End
         Begin VB.TextBox txtAccountHeadCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3405
            TabIndex        =   31
            Top             =   525
            Width           =   1305
         End
         Begin VB.OptionButton optDebit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   285
            TabIndex        =   15
            Top             =   495
            Width           =   1230
         End
         Begin VB.OptionButton optCredit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   285
            TabIndex        =   16
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub.Ledger"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5400
            TabIndex        =   46
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Narration"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1395
            TabIndex        =   32
            Top             =   1170
            Width           =   885
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Account Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2145
            TabIndex        =   30
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   255
            Left            =   30
            TabIndex        =   39
            Top             =   120
            Width           =   10635
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2280
         Left            =   585
         TabIndex        =   33
         Top             =   3735
         Width           =   10710
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8220
            TabIndex        =   36
            Top             =   1965
            Width           =   1815
         End
         Begin VSFlex8LCtl.VSFlexGrid vsGrid 
            Height          =   1470
            Left            =   435
            TabIndex        =   18
            Top             =   480
            Width           =   9885
            _cx             =   17436
            _cy             =   2593
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483644
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmJournalEntry.frx":1E12
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Amount :"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7350
            TabIndex        =   35
            Top             =   1995
            Width           =   825
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   255
            Left            =   15
            TabIndex        =   34
            Top             =   210
            Width           =   10650
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   53
         Top             =   195
         Width           =   405
      End
      Begin VB.Label lblVoucherNo 
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8655
         TabIndex        =   42
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ref No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6525
         TabIndex        =   7
         Top             =   195
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   0
         Top             =   165
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
    Option Explicit
    Dim mTransactionTypeID As Long
    Dim mDataSavedFlag As Boolean
    Dim mChkAmountRP As Variant
    Dim mZonal As Variant 'Added by sunil
    
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearRequestID As Variant
    Public mWebExtractJV As Boolean
    Public mWebExtractJVDate As Date
    
    Public Sub DisplayReceiptDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        Dim mStatus         As Variant
        
        Call FormInitialize
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select tnyStatus From faVouchers"
        mSql = mSql + " Where intVoucherNo = " & mVoucherNo
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), Null, Rec!tnyStatus)
        End If
        Rec.Close
        If mStatus = 0 Or IsNull(mStatus) Then
            mSql = "Select  *,faTransactionChild.tinDebitOrCreditFlag,isNull(faVouchers.intExternalModuleID,0) ModuleID,faSubSidiaryAccountHeads.vchTitle SubLedger,faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID SubLedgerID,faAccountHeads.vchAccountHeadCode AccCode, faTransactions.intFundID FundID,faVouchers.tnyVoucherGroupID tnyVrGroupID  From faVouchers"
            mSql = mSql + " Left Join faTransactions On faTransactions.intVoucherId = faVouchers.intVoucherId" & vbNewLine
            mSql = mSql + " Left Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID " & vbNewLine
            mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID" & vbNewLine
            mSql = mSql + " Left Join faFunctions On fatransactions.intFunctionId = faFunctions.intFunctionId" & vbNewLine
            mSql = mSql + " Left Join faFunctionaries On faTransactions.intFunctionaryId = faFunctionaries.intFunctionaryId" & vbNewLine
            mSql = mSql + " Left Join faFunds On faFunds.intFundId = faTransactions.intFundId" & vbNewLine
            mSql = mSql + " Left Join faFields On faTransactions.intFieldID = faFields.intFieldID" & vbNewLine
    '        mSQL = mSQL + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
            mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID"
            'mSQL = mSQL + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
            mSql = mSql + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
            mSql = mSql + " Left Join faAccountHeads On faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID"
            mSql = mSql + " Left Join faBanks On faVouchers.intKeyID1 = faBanks.intAccountHeadID"
            mSql = mSql + " Left Join faSubSidiaryAccountHeads On faVouchers.numSubLedgerID =faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID"
            mSql = mSql + " Where faVouchers.intVoucherNo = " & mVoucherNo
    '        mSQL = mSQL + " And faVouchers.tnyCancelFlag <> 1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                
                If Rec!dtDate <> gbTransactionDate Then ' AIBY : BLOCKED along with Date Change 09-Oct-2014
                    cmdSave.Enabled = False
                End If
                
                '----------------------------
                If Rec!ModuleID = 55 Then
                    MsgBox "You Are Not Allowed To Edit Reversed Journal Vouchers", vbInformation
                    'Exit Sub
                    cmdSave.Enabled = False
                End If
                '----------------------------
                On Error Resume Next
                cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                If err <> 0 Then
                    MsgBox "Automated JV is not Allowed to Edit", vbInformation
                    cmdSave.Enabled = False
                    'Exit Sub
                Else
                    cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                End If
                
                
                On Error GoTo 0
                
                
                
'                If Not IsNull(Rec!intTransactionTypeID) Then
'                    cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = IIf(IsNull(Rec!intTransactionTypeID), " ", Rec!intTransactionTypeID)
'                End If
                txtVoucherNo.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intVocherID
                txtReference.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                txtReference.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
                'txtFund.Tag = IIf(IsNull(Rec.Fields(34)), "", Rec.Fields(34)) 'intFundID
                txtFund.Tag = IIf(IsNull(Rec!FundID), "", Rec!FundID) 'intFundID
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                txtField.Text = IIf(IsNull(Rec!vchField), "", Rec!vchField)
                txtField.Tag = IIf(IsNull(Rec!intFieldID), "", Rec!intFieldID)
                txtSubsidiaryLedger.Text = IIf(IsNull(Rec!SubLedger), "", Rec!SubLedger)
                txtSubsidiaryLedger.Tag = IIf(IsNull(Rec!SubLedgerID), "", Rec!SubLedgerID)
                
                If Rec!tnyVrGroupID = 2 Then
                    txtRPLink.Text = IIf(IsNull(Rec!numLinkKeyID), "", Rec!numLinkKeyID)
                    chkRP.Value = vbChecked
                    'chkRP.Enabled = True
                    cmdSearchTransactions.Enabled = True
                End If
                
                If Not IsNull(Rec!tinDebitOrCreditFlag) Then
                    If (Rec!tinDebitOrCreditFlag) = 0 Then
                        optDebit.Value = False
                        optCredit.Value = True
                    Else
                        optDebit.Value = True
                        optCredit.Value = False
                    End If
                End If
                'cmbInstruments.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                'cmbInstruments.ItemData(cmbInstruments.ListIndex) = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                txtAccountHeadCode.Text = IIf(IsNull(Rec!AccCode), "", Rec!AccCode)
                txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
                
'                If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
'                    txtAccountNo.Text = IIf(IsNull(Rec!vchAccountNumber), "", Rec!vchAccountNumber)
'                    txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                    dtpIssueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
'                    dtpDueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
'                    txtNameOfBank.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
'                    txtBranch.Text = IIf(IsNull(Rec!vchBranch), "", Rec!vchBranch)
'                End If
                
                txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                If val(txtReference.Tag) <= 0 Then
                    vsGrid.Rows = 1
                    vsGrid.Rows = 15
                Else
                        mSqlAccHeads = "Select * From faTransactionChild"
            '            mSqlAccHeads = mSqlAccHeads + " Inner Join faTransactionChild On faVoucherChild.intAccountHeadID = faTransactionChild.intAccountHeadID"
                        mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                        mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & val(txtReference.Tag)
                        mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
                        RecAccHeads.Open mSqlAccHeads, mCnn
                        mRowCount = 1
                        While Not Rec.EOF
                            While Not RecAccHeads.EOF
                                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
            '                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
                                vsGrid.Rows = vsGrid.Rows + 1
                                mRowCount = mRowCount + 1
                                'If Not IsNull(RecAccHeads!tinDebitOrCreditFlag) Then
                                '    If (RecAccHeads!tinDebitOrCreditFlag) = 0 Then
                                '        optCredit.Value = True
                                '    Else
                                '        optDebit.Value = True
                                '    End If
                                'End If
                                RecAccHeads.MoveNext
                            Wend
                            Rec.MoveNext
                        Wend
                        RecAccHeads.Close
                End If
                Call Calculate
            End If
            Rec.Close
        Else
            MsgBox "Can't edit this entry", vbCritical
            Exit Sub
        End If
    End Sub
       
    Private Sub SaveData(arrInputMaster As Variant, mInput As Variant)
            
            Dim objdb               As New clsDB
            Dim mCnn                As ADODB.Connection
            Dim arrOutPut           As Variant
            Dim arrInput(7)         As Variant
            Dim mintTransactionID   As Long
            Dim mLoop               As Long
            Dim mCount              As Long
            
            objdb.SetConnection mCnn
                'mCnn.BeginTrans
                On Error GoTo ErrRollBack:
                Call objdb.ExecuteSP("spSaveTransactions", arrInputMaster, arrOutPut, , mCnn)
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintTransactionID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack:
                End If
                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mintTransactionID
                For mLoop = 0 To ((UBound(mInput) + 1) / 8) - 1
                    arrInput(0) = mintTransactionID
                    arrInput(1) = mInput(mCount + 1)
                    arrInput(2) = mInput(mCount + 2)
                    arrInput(3) = mInput(mCount + 3)
                    arrInput(4) = mInput(mCount + 4)
                    arrInput(5) = mInput(mCount + 5)
                    arrInput(6) = mInput(mCount + 6)
                    arrInput(7) = mInput(mCount + 7)
                    mCount = mCount + 8
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                Next mLoop
                'mCnn.CommitTrans
                mDataSavedFlag = True
                
            Exit Sub
ErrRollBack:
            Debug.Print Error$
            'mCnn.RollbackTrans
            mDataSavedFlag = False
    End Sub
    
    Private Sub GetDataForUpdation()
        
        Dim objAcc              As New clsAccounts
        Dim objTranType         As New clsTransactionType
        Dim objdb               As New clsDB
        
        Dim mExtCnn             As New ADODB.Connection
        Dim mConStr             As String
        
        Dim Rec                 As New ADODB.Recordset
        Dim RecTranType         As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        
        Dim arrInputMaster      As Variant
        Dim arrInput            As Variant
        Dim mRows               As Long
        Dim mintByLedgerID      As Long
        Dim arrOutPut As Variant
        Dim mLoopCrl As Integer
        
        Dim mintFundID          As Variant
        Dim mintFunctionID      As Variant
        Dim mintFunctionaryID   As Variant
        Dim mintFieldID         As Variant
        Dim mintVoucherID As Variant
        
        Dim mintProcessID       As Long
        Dim mLoop               As Long
        Dim mBudgetCentreID     As Variant
        Dim mtinDebitOrCredit   As Integer
        Dim mAmount             As Double
        Dim mintOrder           As Integer
        Dim mSql                As String
        Dim mVoucherGroupID     As Integer          '' Added By Sinoj On 04 Oct 2009
        Dim mRPLinkID           As Variant
        Dim mDate               As String
        Dim mSubLedgerID        As Variant
        
        Dim mYearID             As Integer
        
        Dim mExternalModuleID As Variant 'Added by Sunil on Aug 22
        Dim numLocationID   As Variant   'Added by Anisha on 02 Jun 2015 For adjustment journal zonalid updation
        Dim mExternalAppID As Variant 'Added by Anisha on Oct 24
        
        If Not IsDate(txtDate.Text) Then
            MsgBox "Plz check the date", vbInformation
            Exit Sub
        End If
        
        '-----------------------------------------------------------'
        ' PREVIOUS YEAR TAKS REQUESTS                               '
        '-----------------------------------------------------------'
        If mPreviousYearMode = 1 Then
            If val(txtAmount.Text) = 0 Then
                If val(txtAmount.Text) <> val(txtAmount.Tag) Then
                    MsgBox "Task Requested Amount was Rs." & Format(val(txtAmount.Tag), "0.00"), vbInformation
                    Exit Sub
                End If
            End If
            mDate = txtDate.Text
            mYearID = gbFinancialYearID - 1
        Else
            mDate = CDate(txtDate.Text)
            mYearID = gbFinancialYearID
        End If
        
        
        mintFundID = Null '-1
        mintFunctionID = Null '-1
        mintFunctionaryID = Null '-1
        mintFieldID = Null '-1
        mBudgetCentreID = Null
        mintProcessID = 0
        '-------------------------------------------------
        If mZonal = 1 Then
            mVoucherGroupID = 5 ' Added by sunil on 22-08-2011
            mExternalModuleID = 45
        Else
            mVoucherGroupID = 0    ' Added By Sinoj On 04 Oct 2009
            mExternalModuleID = 1
        End If
        '------------------------------------------------------
        If mWebExtractJV = True Then
            mExternalAppID = 118
        Else
            mExternalAppID = 115
        End If
        mRPLinkID = ""
        If chkRP.Value = 1 Then
            mVoucherGroupID = 2
            txtNarration.Text = txtNarration.Text & vbNewLine & "The Correction Entry of " & txtRPLink.Text & " Voucher Number"
        End If
            If Trim(txtRPLink.Text) <> "" Then
                mRPLinkID = Trim(txtRPLink.Text)
                If val(txtVoucherNo.Tag) < 1 Then
                    If val(mChkAmountRP) < val(txtAmount.Text) Then
                        MsgBox "The Correction Entry Must be Less than Or Equal to Rs. " & CStr(mChkAmountRP)
                        Exit Sub
                    End If
                End If
            End If                                      '' Added By Sinoj On 04 Oct 2009
            If chkRP.Value = 1 And txtRPLink = "" Then
                MsgBox "The Connection Must be Selected (The Link to the Receipt Entry)"
                Exit Sub
            End If
            On Error Resume Next
            '----------------------------------------------------'
            ' Validations
            '----------------------------------------------------'
            ' Debit Account Head
            If txtAccountHeadCode.Text = "" Then
                MsgBox "Please select AccountHead", vbApplicationModal
                Exit Sub
            End If
            objAcc.SetAccountCode (Trim(txtAccountHeadCode.Text))
            If objAcc.AccountHeadID < 0 Then
                MsgBox "Select a Credit or Debit Account Head!", vbInformation
                txtAccountHeadCode.SetFocus
                Exit Sub
            End If
            '-------------------------'
            ' Debit and Credit Amount '
            '-------------------------'
            Call Calculate
            If val(txtAmount.Text) <= 0 Then
                MsgBox "Check the Amount!!", vbInformation
                vsGrid.SetFocus
                Exit Sub
            End If
            '       If cmbTransactionType.Text = "" Then
            '           MsgBox "Select Transactiontype", vbInformation
            '           Exit Sub
            '       Else
            '           mTransactionTypeID = cmbTransactionType.itemData(cmbTransactionType.ListIndex)
            '       End If
            '       Added By sunil 23-08-2011

            If mZonal = 1 Then
               mTransactionTypeID = frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4)
            Else
                If cmbTransactionType.ListIndex > 0 Then
                       mTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                Else
                      MsgBox "Please Select Transaction Type", vbInformation
                Exit Sub
                End If
            End If
        
            If mTransactionTypeID = 3006 Then ''' Adjustment
               If txtRPLink.Text = "" Then
               'If val(cmdSearchTransactions.Tag) < 1 Then
                   MsgBox "For Adjustment Transaction type.. Please Select Voucher for Adjustment"
                   Exit Sub
               End If
                If val(txtNarration.Tag) <> "" Then
                    numLocationID = val(txtNarration.Tag)
                Else
                    numLocationID = gbLocationID
                End If
            Else
                numLocationID = gbLocationID
               
            End If
        
            
            If val(txtFund.Tag) > 0 Then
                mintFundID = txtFund.Tag
            Else
                MsgBox "Select Fund", vbInformation
                txtFund.SetFocus
                Exit Sub
            End If
                
            If val(txtFunctionary.Tag) > 0 Then
                mintFunctionaryID = txtFunctionary.Tag
            Else
                MsgBox "Select Functionary", vbInformation
                txtFunctionary.SetFocus
                Exit Sub
            End If
            If val(txtFunction.Tag) > 0 Then
                mintFunctionID = txtFunction.Tag
            Else
                MsgBox "Select  Function", vbInformation
                txtFunction.SetFocus
                Exit Sub
            End If
        
        If val(txtField.Tag) > 0 Then
            mintFieldID = txtField.Tag
'        Else
'            MsgBox "Select Field", vbInformation
'            txtField.SetFocus
'            Exit Sub
        End If
        If txtAccountHead.Text <> "" Then
            If optCredit.Value = 0 And optDebit.Value = 0 Then
                MsgBox "Select Credit Or Debit", vbInformation
                Exit Sub
            End If
        End If
        '---Date Change For Editing----------------------------------
        If mPreviousYearMode = 1 Then
            If Not (DateAdd("yyyy", -1, gbStartingDate) <= CDate(mDate) And DateAdd("yyyy", -1, gbEndingDate) >= CDate(mDate)) Then
                MsgBox "Please check the requested TransactionDate is correct ?! ", vbInformation
                Exit Sub
            End If
        Else
            If gDateValidation(CDate(txtDate.Text)) = False Then
                MsgBox "Please Enter Valid Date", vbApplicationModal
                Exit Sub
            End If
        End If
        If txtDate.Text <> "" Then
            mDate = Format(txtDate.Text, "dd/mmm/yy")
        Else
            mDate = gbTransactionDate
        End If
        '-----------------------------------------------------
        
        If txtSubsidiaryLedger.Tag > 0 Then
            mSubLedgerID = IIf(val(txtSubsidiaryLedger.Tag) > 0, val(txtSubsidiaryLedger.Tag), Null)
        End If
        If vsGrid.TextMatrix(vsGrid.Row, 4) <> "" Then
            If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then
                MsgBox "Please select the Account Head", vbCritical
                Exit Sub
            End If
        End If
        If cashBankValidate Then
            MsgBox "Cash/Bank/Treasury Head not Accepted in JV"
            Exit Sub
        End If
        
        On Error GoTo 0
          
            '---------------------------------------------------------------------------------------------'
                'Saving JV details to VOUCHER and VOUCHER CHILD Tables : Added by Aswathi ON 16/08/2008
            '---------------------------------------------------------------------------------------------'
            '-------------------------------------------------------'
                ' faVoucher
            '-------------------------------------------------------'
            
'             @intVoucherID_1     [bigint],
'             @intLocalBodyID_2  [int],
'             @intTransactionID_3    [bigint],
'             @intTransactionTypeID_4    [int],
'             @tnyVoucherTypeID_5    [tinyint],
'             @intVoucherNo_6    [int],
'             @intBookNo_7       [int],
'             @dtDate_8      [smalldatetime],
'             @fltAmount_9       [float],
'             @intInstrumentTypeID_10 [int],
'             @vchInstrumentNo_11    [varchar](50),
'             @dtInstrumentDate_12   [smalldatetime],
'             @vchDescription_13     [varchar](500),
'             @numZoneID_14      [numeric],
'             @numWardID_15      [numeric],
'             @intDoorNoP1_16    [int],
'             @vchDoorNoP2_17    [varchar](10),
'             @vchDoorNoP3_18    [varchar](10),
'             @intUserID_19      [int],
'             @intCounterID_20   [int],
'             @numSubLedgerID_21     [numeric],
'             @intKeyID1_22      [int],
'             @intKeyID2_23      [int],
'             @intExternalApplicationID_24   [int],
'             @intExternalModuleID_25    [int],
'             @intFinancialYearID_26     [int],
'             @tnyShiftID_27     [tinyint] = Null,
'             @tnyPrintFlag_28   [tinyint] = Null,
'             @tnyCancelFlag_29  [tinyint] = Null,
'
'             @vchBank_33    [varchar](50)= Null,
'             @vchBankPlace_34   [varchar](50)= Null,
'             @intFundID_35  [int] = Null

            
            arrInput = Array( _
                                IIf(txtVoucherNo.Tag = "", -1, txtVoucherNo.Tag), _
                                gbLocalBodyID, _
                                Null, _
                                mTransactionTypeID, _
                                40, _
                                Null, _
                                Null, _
                                mDate, _
                                val(txtAmount), _
                                Null, _
                                Null, _
                                Null, _
                                Trim(txtNarration), _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                gbUserID, _
                                gbCounterID, _
                                mSubLedgerID, _
                                val(txtAccountHead.Tag), _
                                Null, mExternalAppID, mExternalModuleID, mYearID, Null, Null, Null, Null, Null, mintFundID, Null, Null, txtReference.Text, Null, Null, Null, Null, numLocationID, mVoucherGroupID, mRPLinkID)
                               
                               
        '-------------------------------------------------------'
        ' Connection And Transaction Begins                     '
        '-------------------------------------------------------'
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'mCnn.BeginTrans                     Commented on 09/052010
        On Error GoTo ErrRollBack:
        
                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID = arrOutPut(0, 0)
                    If mintVoucherID <> "" Then
                        mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        End If
                        Rec.Close
                    End If
                Else
                    GoTo ErrRollBack:
                End If
                
                '-------------------------------------------------------'
                ' faVoucher Child
                '-------------------------------------------------------'
                'Dim mintVoucherID_1         As Double  '
                Dim mintLocalBodyID_2       As Long
                Dim mintSlNo_3              As Long
                Dim mintAccountHeadID_4     As Long
                Dim mtnyDebitOrCredit_5     As Integer
                Dim mintYearID_6            As Long
                Dim mtnyPeriodID_7          As Integer
                Dim mtnyArrearFlag_8        As Integer
                Dim mnumDemandID_9          As Variant
                Dim mfltAmount_10           As Double
                
                mCnn.Execute "Delete From faVoucherChild Where intVoucherID =" & mintVoucherID
                
                For mLoopCrl = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpText, mLoopCrl, 1) <> "" Then
                        objAcc.SetAccountCode (vsGrid.Cell(flexcpText, mLoopCrl, 1))
                        mintLocalBodyID_2 = gbLocalBodyID
                        mintSlNo_3 = mLoopCrl
                        
                        mintAccountHeadID_4 = objAcc.AccountHeadID
 '                       mtnyDebitOrCredit_5 = 0
                        mintYearID_6 = mYearID
                        mtnyPeriodID_7 = 3
                        mtnyArrearFlag_8 = 0
                        '------Added by Sunil -----------
                        If mZonal = 1 Then
                            mnumDemandID_9 = frmTransactionTypeWiseDemandInbox.txtDemandNo.Tag
                        Else
                            mnumDemandID_9 = Null
                        End If
                        '--------------------------------
                        mfltAmount_10 = val(vsGrid.Cell(flexcpText, mLoopCrl, 4))
                        If optCredit.Value = True Then      '  According to Option button
                            mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
                        Else                                '  0 = Credit  and 1 = Debit
                            mtinDebitOrCredit = 1           '-------------------------------------'
                        End If
                        '------------------------------------------------'
                            'faVoucherChild Parameters
                        '------------------------------------------------'
                        
'                        @intVoucherID_1     [bigint],
'                        @intLocalBodyID_2  [int],
'                        @intSlNo_3     [int],
'                        @intAccountHeadID_4    [int],
'                        @tnyDebitOrCredit_5    [tinyint],
'                        @intYearID_6   [int],
'                        @tnyPeriodID_7     [tinyint],
'                        @tnyArrearFlag_8   [tinyint],
'                        @numDemandID_9     [numeric],
'                        @fltAmount_10      [float] = 0
                        
                        arrInput = Array( _
                                            mintVoucherID, _
                                            mintLocalBodyID_2, _
                                            mintSlNo_3, _
                                            mintAccountHeadID_4, _
                                            IIf(mtinDebitOrCredit = 0, 1, 0), _
                                            mintYearID_6, _
                                            mtnyPeriodID_7, _
                                            mtnyArrearFlag_8, _
                                            mnumDemandID_9, _
                                            mfltAmount_10 _
                                            )
                        objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                    Else
                        Exit For
                    End If
                Next mLoopCrl
                '-------------------------------------------------------'
                ' faVoucher Address
                '-------------------------------------------------------'
                                                                                                    
                                                                                                    
                                                                                                    
                                                                                                    
                                                                                                    
                                                                                                    
                                                                                                    
                                                                                                    
                '                Else
                '----------------------------------------------------------------'
                '         G e n e r a l - J o u r n a l  P o s t i n g           '
                '----------------------------------------------------------------'
                
                    'Dim mExternalModuleID As Variant    'Added by sunil 22-08-2011
                    If mZonal = 1 Then
                        mExternalModuleID = 45
                    Else
                        mExternalModuleID = 0
                    End If
                    
                    If optCredit.Value = True Then      '  According to Option button
                        mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
                    Else                                '  0 = Credit  and 1 = Debit
                        mtinDebitOrCredit = 1           '-------------------------------------'
                    End If
                    mintProcessID = 0                   ' Used for Automation - Recurring Process
                    
                    If Trim(txtBudgetCentreCode.Text) <> "" Then
                        Dim objBgt As New clsBudgetCentre
                        objBgt.SetBudgetCentre Trim(txtBudgetCentreCode)
                        mBudgetCentreID = objBgt.BudgetCentreID
                    Else
                        mBudgetCentreID = Null
                        'GoTo ErrHandle
                    End If
                    '-------------------------------------'
                    ' Data for Transaction Table          '
                    '-------------------------------------'
                    arrInputMaster = Array( _
                                            IIf(txtReference.Tag = "", -1, txtReference.Tag), _
                                            gbLocalBodyID, _
                                            mYearID, _
                                            Format(mDate, "DD/MMM/YYYY"), _
                                            0, _
                                            mExternalModuleID, _
                                            mintFunctionID, _
                                            mintFunctionaryID, _
                                            mintFieldID, _
                                            mintFundID, _
                                            mBudgetCentreID, _
                                            txtNarration.Text, _
                                            400, _
                                            mintProcessID, _
                                            "JV", _
                                            40, _
                                            Null, _
                                            IIf(val(txtSubsidiaryLedger.Tag) > 0, val(txtSubsidiaryLedger.Tag), Null), _
                                            gbUserID, _
                                            mintVoucherID, _
                                            mVoucherGroupID)                    '' Added Sinoj On 04 Oct 2009

                                            
                        '----------------------------------------'
                        ' Data for TransactionChild First Record '
                        '----------------------------------------'
                        objAcc.SetAccountCode Trim(txtAccountHeadCode.Text)
                        If objAcc.AccountHeadID = -1 Then
                            GoTo ErrRollBack:
                        Else
                            mintByLedgerID = objAcc.AccountHeadID
                        End If
                        mintOrder = 1
                        arrInput = Array(-1, _
                                        mintOrder, _
                                        mintByLedgerID, _
                                        Format(val(txtAmount.Text), "0.00"), _
                                        mtinDebitOrCredit, _
                                        "", _
                                        "", _
                                        mintFundID _
                                        )
                    '----------------------------------------'
                    ' Data for TransactionChild From Grid    '
                    '----------------------------------------'
                    For mLoop = 1 To vsGrid.Rows - 1
                            If vsGrid.TextMatrix(mLoop, 1) <> "" Then
                                objAcc.SetAccountCode (Trim(vsGrid.TextMatrix(mLoop, 1)))
                                If objAcc.AccountHeadID = -1 Then
                                    GoTo ErrChk:
                                End If
                            Else
ErrChk:
                                If val(txtAmount) = mAmount Then
                                    Exit For
                                Else
                                    GoTo ErrRollBack
                                End If
                            End If
                            
                            ReDim Preserve arrInput(UBound(arrInput) + 8)
                            mintOrder = mintOrder + 1
                            arrInput(UBound(arrInput) - 7) = -1
                            arrInput(UBound(arrInput) - 6) = mintOrder
                            arrInput(UBound(arrInput) - 5) = objAcc.AccountHeadID
                            arrInput(UBound(arrInput) - 4) = Format(val(vsGrid.TextMatrix(mLoop, 4)), "0.00")
                            arrInput(UBound(arrInput) - 3) = IIf(mtinDebitOrCredit = 0, 1, 0)
                            arrInput(UBound(arrInput) - 2) = mintByLedgerID
                            arrInput(UBound(arrInput) - 1) = Trim(vsGrid.TextMatrix(mLoop, 3))
                            arrInput(UBound(arrInput)) = mintFundID
                            mAmount = mAmount + Format(val(vsGrid.TextMatrix(mLoop, 4)), "0.00")
                    Next mLoop
                    If val(txtAmount) <> mAmount Then
                        GoTo ErrRollBack
                    End If
                    Call SaveData(arrInputMaster, arrInput)
                     
                     
                     
                     '==============For Zonal Collection Update DemandStatus in HO=================By sunil on 22-08-2011
                    If mZonal = 1 Then
                      Dim mCnnHO As New ADODB.Connection
                       If (objdb.CreateNewConnection(mCnnHO, enuSourceString.SaankhyaHO)) Then
                                arrInput = Array(mnumDemandID_9, mintVoucherID, gbTransactionDate, 11, mTransactionTypeID)
                                objdb.ExecuteSP "spUpdateDemandChildStatus", arrInput, , , mCnnHO, adCmdStoredProc
                            Else
                                MsgBox "SaankhyaHo Connection Does not exists"
                      End If
                   End If
                    '====================================================================================
                    If mDataSavedFlag = False Then
                        GoTo ErrRollBack
                    End If
                    ' mCnn.CommitTrans     Commented on 09/0520104
                   
                   
                    '--------------------------------------------------------------------------'
                    ' PREVIOUS YEAR TASK REQUEST
                    '--------------------------------------------------------------------------'
                    If mPreviousYearMode = 1 Then
                        mSql = "Update faPendingTaskRequest SET tnyStatus = 8 WHERE intRequestID = " & mPreviousYearRequestID
                        mCnn.Execute mSql
                    End If
                    If mWebExtractJV = True Then
                        mSql = "Update faWebExtracts SET numKeyID = " & mintVoucherID & " WHERE intwebExtractID = " & txtBudgetCentreCode.Tag
                        mCnn.Execute mSql
                    End If
                    
                   
                   
            Exit Sub
ErrRollBack:
       ' mCnn.RollbackTrans               Commented on 09/052010
        mDataSavedFlag = False
    End Sub
    
    Private Sub FillGridCombo()
            Dim objdb As New clsDB
            Dim RecAccHead As New ADODB.Recordset
            Dim mItem As String
            
            RecAccHead.CursorLocation = adUseClient
            Set RecAccHead = GetRecordSet("spGetAccHead4Fill", adOpenStatic, adLockReadOnly)
            While Not RecAccHead.EOF
                mItem = mItem + "|" + RecAccHead!vchAccountHead
                RecAccHead.MoveNext
            Wend
            RecAccHead.Close
            vsGrid.ColComboList(2) = mItem
    End Sub
    
    Private Sub GetTransactionTypeMapping(mExtTranID As Long)
        
        '------------------------------------------------------------'
        '                                                            '
        '------------------------------------------------------------'
        ' 1) Identify External Data source
        ' 2) Get TransactionType
        ' 3) Map with External AccountHead Masters with AccountHeads
        ' 4) Get Data ( Amount) and Populate vsGrid
        '------------------------------------------------------------'
        ' mExtTranID => PrimaryID of Transaction table of External
        '               Application
        '------------------------------------------------------------'
        
        Dim objTranType As New clsTransactionType
        Dim objdb As New clsDB
        
        Dim mExtCnn As New ADODB.Connection
        Dim mConStr As String
        
        Dim Rec As New ADODB.Recordset
        Dim RecTranType As New ADODB.Recordset
        
        Dim arrInput As Variant
        Dim mRows As Long
        
        objTranType.SetTransactionType (mTransactionTypeID)
        Select Case objTranType.ExternalApplicationID
            Case Is = AppID.Payroll
                
                '---------------------------------------------'
                ' DB_faExternalTransactions
                '---------------------------------------------'
                arrInput = Array(mTransactionTypeID)
                mConStr = objdb.GetConnectionString(2)
                objdb.SetExtDBConnection mExtCnn, mConStr
                
                
                'Set RecTranType = objDB.ExecuteSP("spGetTranTypeDetailsByHeads", arrInput)
                Set Rec = GetRecordSet("spGetExtTransactionDetails " & mExtTranID, , , mExtCnn)
                'Set Rec = objDB.ExecuteSP("spGetExtTransactionDetails", arrInput, , False, mExtCnn)
                Set RecTranType = GetRecordSet("spGetTranTypeDetailsByHeads " & mTransactionTypeID, adOpenDynamic)
                
                vsGrid.Rows = 1
                                
                If Not (RecTranType.BOF And RecTranType.EOF) Then
                    While Not RecTranType.EOF
                        If RecTranType!intOrder = 1 Then
                            txtAccountHeadCode.Text = RecTranType!vchAccountHeadCode
                            txtAccountHead.Text = RecTranType!vchAccountHead
                            optDebit.Value = RecTranType!tinDebitOrCredit
                        End If
                        
                        Rec.MoveFirst
                        While Not Rec.EOF
                            If Rec!intExtAccountHeadID = RecTranType!intExternalAccountHeadID Then
                                mRows = mRows + 1
                                vsGrid.AddItem mRows & vbTab & RecTranType!vchAccountHeadCode & vbTab & RecTranType!vchAccountHead & vbTab & Rec!vchNarration & vbTab & Format(Rec!floAmount, "0.00"), mRows
                            End If
                            Rec.MoveNext
                        Wend
                        RecTranType.MoveNext
                        
                    Wend
                End If
                
                RecTranType.Close
                Rec.Close
                Set mExtCnn = Nothing
        End Select
    
    
    End Sub
    
    
    Private Sub FormInitialize()
        cmbTransactionType.ListIndex = -1
        txtReference.Text = ""
        txtBudgetCentreCode.Text = ""
        lstMasters.ListIndex = -1
        txtAccountHeadCode.Text = ""
        txtAccountHead.Text = ""
        txtNarration.Text = ""
        txtAmount.Text = ""
        txtVoucherNo.Tag = ""
        
        '''..........Added On 20 Jun 2015
        txtVoucherNo.Text = ""
        txtRPLink.Text = ""
        txtRPLink.Tag = ""
        '''..........
        
        fraBudget.Enabled = True
        
        vsGrid.Visible = False
        vsGrid.Rows = 1
        vsGrid.Rows = 50
        vsGrid.Visible = True
        mDataSavedFlag = False
        mTransactionTypeID = -1
        
        txtSubsidiaryLedger.Text = ""
        txtSubsidiaryLedger.Tag = ""
        txtClaiment = ""
        
        txtDate.Text = gbTransactionDate
        txtDate.Locked = True
        
    End Sub
    
    Private Sub ClearGridAndHeads()
        txtReference.Text = ""
        
        txtAccountHeadCode.Text = ""
        txtAccountHead.Text = ""
        txtNarration.Text = ""
        txtAmount.Text = ""
        
        vsGrid.Visible = False
        vsGrid.Rows = 1
        vsGrid.Rows = 50
        vsGrid.Visible = True
    End Sub
    Private Sub FetchTransactionTypeDetails()
        
        Dim objdb As New clsDB
        Dim arrInput As Variant
        Dim Rec As New ADODB.Recordset
        arrInput = Array(Trim(cmbTransactionType.Text))
        
        Set Rec = objdb.ExecuteSP("spGetTransactionTypeDetails", arrInput, , False)
        If Not (Rec.BOF And Rec.EOF) Then
            
            If IsNull(Rec!vchBudgetCentreCode) Then
                txtBudgetCentreCode.Text = ""
            Else
                txtBudgetCentreCode.Text = Rec!vchBudgetCentreCode
            End If
            
            If IsNull(Rec!vchFunctionary) Then
                txtFunctionary.Text = ""
            Else
                txtFunctionary.Text = Rec!vchFunctionary
                txtFunctionary.Tag = Rec!intFunctionaryID
            End If
            
            If IsNull(Rec!vchFunction) Then
                txtFunction.Text = ""
            Else
                txtFunction.Text = Rec!vchFunction
                txtFunction.Tag = Rec!intFunctionID
            End If
            
            If IsNull(Rec!vchField) Then
                txtField.Text = ""
            Else
                txtField.Text = Rec!vchField
                txtField.Tag = Rec!intFieldID
            End If
            
            If IsNull(Rec!vchFund) Then
                txtFund.Text = ""
            Else
                txtFund.Text = Rec!vchFund
                txtFund.Tag = Rec!intFundID
            End If
            'fraBudget.Enabled = False
        Else
            txtBudgetCentreCode.Text = ""
            txtFunctionary.Text = ""
            txtFunction.Text = ""
            txtField.Text = ""
            txtFund.Text = ""
            fraBudget.Enabled = True
        End If
        
    End Sub
    
    
      Private Sub Calculate()
            Dim mLoopCount As Long
            Dim mAmt As Currency
            For mLoopCount = 1 To vsGrid.Rows - 1
                If val(vsGrid.TextMatrix(mLoopCount, 4)) > 0 Then
                    vsGrid.TextMatrix(mLoopCount, 4) = Format(val(vsGrid.TextMatrix(mLoopCount, 4)), "0.00")
                    mAmt = mAmt + val(vsGrid.TextMatrix(mLoopCount, 4))
                Else
                    Exit For
                End If
            Next mLoopCount
            txtAmount.Text = Format(mAmt, "0.00")
            
    End Sub
      
    Private Sub chkRP_Click()
        If val(txtVoucherNo.Tag) < 1 Then
            If chkRP.Value = vbChecked Then
                cmdSearchTransactions.Enabled = True
                Call cmdSearchTransactions_Click '' Seach Transaction Calling
            Else
                cmdSearchTransactions.Enabled = False
                txtRPLink.Text = ""
                txtRPLink.Tag = 0
                cmdSearchTransactions.Tag = 0
                txtAccountHeadCode.Text = ""
                txtAccountHead.Text = ""
                txtAccountHead.Tag = 0
                
            End If
        Else
            If (chkRP.Value) = vbUnchecked Then
                cmdSearchTransactions.Enabled = True
                txtRPLink.Text = ""
                txtRPLink.Tag = 0
                cmdSearchTransactions.Tag = 0
                txtAccountHeadCode.Text = ""
                txtAccountHead.Text = ""
                txtAccountHead.Tag = 0
                vsGrid.Clear 0, 1
                txtNarration.Text = ""
                
            End If
        End If
    End Sub



    Private Sub cmbTransactionType_Click()
        Call FetchTransactionTypeDetails
        If cmbTransactionType.ListIndex > 0 Then
            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 3006 Then
                fraAdjustments.Enabled = True
            Else
                chkRP.Value = 0
                txtRPLink.Text = ""
                txtRPLink.Tag = -1
                fraAdjustments.Enabled = False
            End If
        End If
    End Sub

    Private Sub cmbTransactionType_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub cmbTransactionType_LostFocus()
        Dim mTempID As Long
        mTempID = mTransactionTypeID
        
        If cmbTransactionType.ListIndex > 0 Then
            mTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
        Else
            mTransactionTypeID = -1
        End If
        
        If txtReference.Text <> "" Then
            If mTempID <> mTransactionTypeID Then
                Call FormInitialize
            End If
        End If
    End Sub
    Private Sub cmdCancel_Click()
'        Call GetDataForUpdation
        Unload Me
    End Sub

    Private Sub cmdField_Click()
        Dim mSql    As String
        
        mSql = "Select vchField,intFieldID From faFields Order By vchField"
        Call PopulateList(lstMasters, mSql, , True, , True)
        lstMasters.Tag = "3"
        lstMasters.Visible = True
        lstMasters.SetFocus
    End Sub
    
    Private Sub cmdFunction_Click()
            Dim mSql As String
            Dim mToken1 As String
            
            frmSearchFunction.Show vbModal         'Modified
            mToken1 = Token(gbSearchStr, " ")
            txtFunction.Text = Trim(gbSearchStr)
            txtFunction.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            
            'mSql = "Select vchFunction, intFunctionID From faFunctions Order By vchFunction"
            'Call PopulateList(lstMasters, mSql, , True, , True)
            'lstMasters.Tag = "1"
            'lstMasters.Visible = True
            'lstMasters.SetFocus
    End Sub
    


    Private Sub cmdFunctionary_Click()
            Dim mSql As String
            Dim mToken1 As String
            frmSearchFunctionary.Show vbModal       'Modified
            mToken1 = Token(gbSearchStr, " ")
            txtFunctionary.Text = Trim(gbSearchStr)
            txtFunctionary.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            
            'mSql = "Select vchFunctionary, intFunctionaryID From faFunctionaries Order By vchFunctionary"
            'Call PopulateList(lstMasters, mSql, , True, , True)
            'lstMasters.Tag = "2"
            'lstMasters.Visible = True
            'lstMasters.SetFocus
    End Sub

    Private Sub cmdFund_Click()
        Dim mSql As String
        mSql = "Select vchFund, intFundID From faFunds Where tnyActiveFlag = 1 Order By vchFund"
        Call PopulateList(lstMasters, mSql, , True, , True)
        lstMasters.Tag = "4"
        lstMasters.Visible = True
        lstMasters.SetFocus
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdSave.Enabled = True
        txtVoucherNo.Text = ""
        txtVoucherNo.Tag = ""
        chkRP.Value = vbUnchecked
        txtRPLink.Text = ""
        txtRPLink.Tag = ""
    End Sub

    Private Sub cmdSave_Click()
        Call GetDataForUpdation
        If mDataSavedFlag Then
'            Call FormInitialize
            cmdSave.Enabled = False
            mPreviousYearMode = 0
            mPreviousYearRequestID = Null
        End If
    End Sub

    Private Sub cmdSearchAccountHead_Click()
        Dim mSql As String
'        If cmbInstruments.ListIndex > 0 Then
'            Select Case cmbInstruments.ItemData(cmbInstruments.ListIndex)
'                Case 1 '[Cash]
'                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.tinHiddenFlag = 0 AND  faAccountHeads.intGroupID = " & faCash
                'Case 7 '[Treasury Bills]
                ' mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
'                Case Else
'                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID =" & faBank
'            End Select
            If Trim(txtRPLink.Text) = "" Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where intGroupID Is Null And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
            Else
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where intAccountHeadID in(Select intAccountHeadID From faVoucherChild Where intVoucherID = " & txtRPLink.Tag & ") Order by vchAccountHeadCode"
            End If
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.VoucherMode = 400
            frmSearchAccountHeads.Show vbModal
            txtAccountHeadCode.SetFocus
'        End If
    End Sub
    Private Sub GetFunctionFunctionary(intVoucherID As Double)
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
       
        mSql = " SELECT intVoucherID, faTransactions.intFunctionID, vchFunctionCode, vchFunction,vchSourceFundName,intSourceFundID,"
        mSql = mSql + " faFunctionaryFunctions.intFunctionaryID , vchFunctionaryCode, vchFunctionary"
        mSql = mSql + " From faTransactions"
        mSql = mSql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faTransactions.intFunctionID"
        mSql = mSql + " INNER JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
        mSql = mSql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID"
        mSql = mSql + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faTransactions.intFundID"
        mSql = mSql + " Where intVoucherID = " & intVoucherID
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.BOF And Rec.EOF) Then
                txtFunction.Text = Rec!vchFunction
                txtFunction.Tag = Rec!intFunctionID
                'txtFund.Text = Rec!vchSourceFundName
                'txtFund.Tag = Rec!intSourceFundID
                txtFunctionary.Text = Rec!vchFunctionary
                txtFunctionary.Tag = Rec!intFunctionaryID
            End If
            Rec.Close
        End If
        
        
    End Sub
    Private Sub cmdSearchTransactions_Click()
        frmSearchTransactions.Receipt = True
        frmSearchTransactions.fmeGroup.Enabled = True
        frmSearchTransactions.Show vbModal
        If gbSearchID > 0 Then
            cmdSearchTransactions.Tag = gbSearchID
            txtRPLink.Tag = gbSearchID
            txtRPLink.Text = gbSearchStr
            mChkAmountRP = gbSearchCode
            Call AdjustMentCheck(val(cmdSearchTransactions.Tag))
            Call GetFunctionFunctionary(gbSearchID)
            
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = "" '' Uses as Amount
            
        End If
    End Sub
    
    Private Function AdjustMentCheck(ByVal VrID As Double)
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim numZonal    As Double
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faVouchers Where intVoucherID=" & VrID
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.BOF And Rec.EOF) Then
                numZonal = IIf(IsNull(Rec!numLocationID), 0, Rec!numLocationID)
                txtNarration.Tag = numZonal
                If IIf(IsNull(Rec!tnyReversed), 0, Rec!tnyReversed) = 1 Then
                    MsgBox "This Voucher is Reversed .. can't Do Adjustment Entry.."
                    txtRPLink.Tag = ""
                    txtRPLink.Text = ""
                    mChkAmountRP = ""
                    Exit Function
                End If
               '' Commented on 11 Mar 2019 suggested by Jiju Krishnam (DE) to allow adjustment for E bill Vouchers
'                If IIf(IsNull(Rec!intExternalApplicationID), 0, Rec!intExternalApplicationID) = 118 Then
'                    MsgBox "This is E bill Generated Voucher .. can't Do Adjustment Entry.."
'                    txtRPLink.Tag = ""
'                    txtRPLink.Text = ""
'                    mChkAmountRP = ""
'                    Exit Function
'                End If
                
                If IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus) = 4 Then
                    MsgBox "This Voucher is Cancelled .. can't Do Adjustment Entry.."
                    txtRPLink.Tag = ""
                    txtRPLink.Text = ""
                    mChkAmountRP = ""
                    Exit Function
                End If
            End If
            mSql = "Select * From faVouchers Where numLinkKeyID in (Select intVoucherID From faVouchers Where intVoucherNo=" & val(txtRPLink.Text) & ")"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.BOF And Rec.EOF) Then
                    MsgBox "This Voucher has AdjustmentEntry .. can't Do Another Adjustment Entry .."
                    txtRPLink.Tag = ""
                    txtRPLink.Text = ""
                    mChkAmountRP = ""
                    Exit Function
            End If
        End If
    End Function
    Private Sub cmdSearchVoucherNo_Click()
        frmSearchJournalVouchers.Show vbModal
        txtVoucherNo.Text = gbSearchStr
        txtVoucherNo.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        If val(txtVoucherNo.Tag) Then
            Call txtVoucherNo_LostFocus
        Else
            txtVoucherNo.SetFocus
        End If
    End Sub

    Private Sub cmdSubLedger_Click()
        Dim objSubLedger As New clsSubLedger
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchID = -1 Then Exit Sub
        txtSubsidiaryLedger.Text = gbSearchStr
        txtSubsidiaryLedger.Tag = gbSearchID
        objSubLedger.SetSubLedgerDetails (gbSearchID)
        txtClaiment.Visible = True
        txtClaiment.Text = objSubLedger.HouseOrOffice
        txtClaiment.Text = txtClaiment.Text + vbNewLine + objSubLedger.LocalPlace
        txtClaiment.Text = txtClaiment.Text + vbNewLine + objSubLedger.MainPlace
        txtClaiment.Text = txtClaiment.Text + vbNewLine + objSubLedger.Street
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub
    Private Sub Form_Load()
        
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        If mWebExtractJV = True Then
            mSql = "Select  replace(vchTransactionType, 'Development Project Expenditure', 'Dev Pr Exp') as vchTransactionType, intTransactionTypeID From faTransactionType Where intTransactionTypeID in  (1141,1151,1161,1171,1181,1191) AND IsNull(tnyHidden,0) <> 1 Order By vchTransactionType"
        Else
            mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID=40 AND IsNull(tnyHidden,0) <> 1 Order By vchTransactionType"
        
        End If
        cmbTransactionType.ToolTipText = cmbTransactionType
        PopulateList cmbTransactionType, mSql, , True, True, True
        
        txtBudgetCentreCode.Visible = False
        Label4.Visible = False
        lstMasters.Visible = False
'        Call FillGridCombo
        vsGrid.ColComboList(1) = "|..."
        Call FormInitialize
        
        '-----------------------------------------------------------------------'
        ' SET PREVIOUS YEAR MODE                                                '
        '-----------------------------------------------------------------------'
        If mPreviousYearMode = 1 Then
            If objdb.SetConnection(mCnn) Then
                mSql = "SELECT * FROM faPendingTaskRequest WHERE intRequestID= " & mPreviousYearRequestID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    
                    txtDate.Text = DdMmmYy(Rec!dtTransactionDate)
                    txtDate.Enabled = False
                    txtAmount.Tag = Rec!fltAmount
                    
                    
                End If
            End If
        End If
        
        
    End Sub

    Private Sub lstMasters_DblClick()
        If lstMasters.ListIndex > -1 Then
            gbSearchStr = lstMasters.Text
            gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
            Select Case val(lstMasters.Tag)
                Case 1: txtFunction.SetFocus
                Case 2: txtFunctionary.SetFocus
                Case 3: txtField.SetFocus
                Case 4: txtFund.SetFocus
            End Select
        End If
    End Sub

    Private Sub lstMasters_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
            Call lstMasters_DblClick
        End If
    End Sub

    Private Sub lstMasters_LostFocus()
        If lstMasters.ListIndex > -1 Then
            gbSearchStr = lstMasters.Text
            gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
        End If
        lstMasters.Visible = False
        Select Case val(lstMasters.Tag)
            Case 1: txtFunction.SetFocus
            Case 2: txtFunctionary.SetFocus
            Case 3: txtField.SetFocus
            Case 4: txtFund.SetFocus
        End Select
    End Sub

    Private Sub optCredit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub optDebit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub

    Private Sub txtAccountHeadCode_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            Dim mCnn As New ADODB.Connection
            Dim objdb As New clsDB
            Dim mSPOut As Variant
            txtAccountHeadCode.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            ''Newly Added By sinoj''
            If Trim(txtRPLink.Text) <> "" And txtAccountHead.Text <> "" Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                objdb.ExecuteSP "Select Max(fltAmount) From faVoucherChild Where intVoucherID = " & txtRPLink.Tag & " And intAccountHeadID = " & txtAccountHead.Tag, , mSPOut, , mCnn, adCmdText
                mChkAmountRP = mSPOut(0, 0)
            End If
            ''----------------------''
            gbSearchStr = ""
            gbSearchID = -1
        End If
        txtAccountHeadCode.SelStart = 0
        txtAccountHeadCode.SelLength = Len(txtAccountHeadCode)
    End Sub

    Private Sub txtAccountHeadCode_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads WHERE  tinHiddenFlag = 0 "
            frmSearchAccountHeads.Show vbModal
            txtAccountHeadCode.SetFocus
        End If
    End Sub

    Private Sub txtAccountHeadCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtBudgetCentreCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
        
    Private Sub txtBudgetCentreCode_LostFocus()

        '------------------------------------------------------'
        ' Searches and Finding Function, Functionary and Field '
        '------------------------------------------------------'
        txtBudgetCentreCode = Trim(txtBudgetCentreCode)
        If Len(Trim(txtBudgetCentreCode)) Then
            Dim objBudCen As New clsBudgetCentre
            objBudCen.SetBudgetCentre (txtBudgetCentreCode.Text)
            If objBudCen.BudgetCentreID < 0 Then
                txtBudgetCentreCode.Text = ""
                txtFunction.Text = ""
                txtFunctionary.Text = ""
                txtField.Text = ""
                txtFund.Text = ""
            Else
                txtFund.Text = objBudCen.FundName
                txtFunction.Text = objBudCen.FunctionName
                txtFunctionary.Text = objBudCen.FunctionaryName
                txtField.Text = IIf(objBudCen.FieldName = "", " ", objBudCen.FieldName)
            End If
        End If
        
    End Sub

    Private Sub txtDate_DblClick()
'''       If txtVoucherNo.Tag <> "" Then
'''            If MsgBox("Do You want to Change the date", vbYesNo) = vbYes Then
'''                txtDate.Locked = False
'''                txtDate.SetFocus
'''            Else
'''                txtDate.Locked = True
'''            End If
'''        Else
'''            Exit Sub
'''        End If
    End Sub

    Private Sub txtDate_LostFocus()
        txtDate.Text = CheckDateInMMM(txtDate.Text)
        If gDateValidation(CDate(txtDate.Text)) = False Then
            MsgBox "Please Enter Valid Date", vbApplicationModal
            Exit Sub
        End If
    End Sub
'    Private Function DateValidation() As Boolean
'        Dim mSql    As String
'        Dim Rec     As New ADODB.Recordset
'        Dim mCnn    As New ADODB.Connection
'        Dim objDb   As New clsDb
'        mSql = "Select  top 1 dtStartingDate From faFinancialYear Order by intFinancialYear Asc"
'        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            If Not (txtDate.Text > CheckDateInMMM(Rec!dtStartingDate) And txtDate.Text <= CheckDateInMMM(gbTransactionDate)) Then
'                DateValidation = False
'            Else
'                DateValidation = True
'            End If
'        End If
'    End Function
    Private Sub txtReference_GotFocus()
        If Len(gbSearchStr) Then
            txtReference.Text = gbSearchStr
            gbSearchStr = ""
            txtReference.SelStart = 0
            txtReference.SelLength = Len(txtReference)
        End If
    End Sub

    Private Sub txtReference_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            If cmbTransactionType.ListIndex > 0 Then
                frmSearchExternalData.TransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                frmSearchExternalData.Show vbModal
            Else
                On Error Resume Next
                cmbTransactionType.SetFocus
                On Error GoTo 0
            End If
        End If
    End Sub
    
    Private Sub txtReference_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
        
    Private Sub txtReference_LostFocus()
        txtReference = Trim(txtReference)
        If Len(txtReference) Then
            If mTransactionTypeID > -1 Then
                
                    Dim objdb As New clsDB
                    Dim mExtCnn As New ADODB.Connection
                    Dim mConStr As String
                    Dim Rec As New ADODB.Recordset
                    Dim arrInput As Variant
                    Dim arrOutPut As Variant
                    Dim objTranType As New clsTransactionType
                    
                    objTranType.SetTransactionType (mTransactionTypeID)
                    Select Case objTranType.ExternalApplicationID
                        Case Is = AppID.Payroll
                            arrInput = Array(mTransactionTypeID)
                            mConStr = objdb.GetConnectionString(2)
                            objdb.SetExtDBConnection mExtCnn, mConStr
                            arrInput = Array(mTransactionTypeID, txtReference.Text)
                            
                            Set Rec = objdb.ExecuteSP("spGetExtTransactionID", arrInput, , False, mExtCnn)
                            If Not (Rec.BOF And Rec.EOF) Then
                                txtReference.Text = Rec!vchExtTransactionCode
                                Call GetTransactionTypeMapping(Rec!intExtTransactionID)
                            Else
                                Call ClearGridAndHeads
                            End If
                        
                    End Select
            End If
        Else
            Call ClearGridAndHeads
        End If
        Call Calculate
    End Sub
    
    Private Sub txtField_GotFocus()
        If gbSearchStr <> "" Then
            txtField.Text = gbSearchStr
            txtField.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtFunction_GotFocus()
        If gbSearchStr <> "" Then
            txtFunction.Text = gbSearchStr
            txtFunction.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtFunctionary_GotFocus()
        If gbSearchStr <> "" Then
            txtFunctionary.Text = gbSearchStr
            txtFunctionary.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtFund_GotFocus()
        If gbSearchStr <> "" Then
            txtFund.Text = gbSearchStr
            txtFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtNarration_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            If Len(txtNarration) Then
                If Asc(Right(txtNarration.Text, 1)) = 10 Then
                    Call PressTabKey
                End If
            End If
        End If
    End Sub
    


    Private Sub txtVoucherNo_LostFocus()
        If txtVoucherNo.Text <> "" Then
            If IsNumeric(txtVoucherNo.Text) Then
                If mID(Trim(txtVoucherNo.Text), 1, 1) <> "4" Then
                    MsgBox "Invalid Journal Voucher Number", vbInformation
                    Exit Sub
                End If
                Call DisplayReceiptDetails(txtVoucherNo.Text)
            Else
                MsgBox "Please Enter Valid Journal Voucher No", vbInformation
            End If
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGrid.Row > 1 Then
            If vsGrid.TextMatrix(vsGrid.Row - 1, 1) = "" Or _
                vsGrid.TextMatrix(vsGrid.Row - 1, 2) = "" Or _
                val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 Then
                Cancel = True
                Exit Sub
            End If
        End If
        If Len(gbSearchStr) Then
            vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
            vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
            vsGrid.Col = vsGrid.Col + 2
            vsGrid.Redraw = flexRDDirect
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinHiddenFlag <> 1 And intGroupID is Null Order by vchAccountHeadCode"
        frmSearchAccountHeads.VoucherMode = 401
        frmSearchAccountHeads.Show vbModal
    End Sub
    
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        If Col = 2 And Trim(vsGrid.Text) <> "" Then
            Set objAccHead = New clsAccounts
            If objAccHead.FindAccountByHead(Trim(vsGrid.Text)) Then
                vsGrid.TextMatrix(vsGrid.Row, 1) = objAccHead.AccountCode
            End If
        ElseIf Col = 4 Then
            vsGrid.TextMatrix(Row, 4) = Format(val(vsGrid.TextMatrix(Row, 4)), "0.00")
            Call Calculate
        End If
    End Sub
    
    Private Sub vsGrid_Validate(Cancel As Boolean)
        If vsGrid.Col = 3 Then
            If Len(vsGrid.TextMatrix(vsGrid.Row, 3)) > 100 Then
                vsGrid.TextMatrix(vsGrid.Row, 3) = Left(vsGrid.TextMatrix(vsGrid.Row, 3), 100)
            End If
        End If
    End Sub
   Public Property Let ZonalCollection(mData As Integer) 'Added by Sunil Babu
        mZonal = mData
    End Property
    Private Function cashBankValidate() As Boolean
    ''' Codded ON 4.7.12 By Anisha
    ''' Validation of Blocking Cash/Bank Heads For Jv
        Dim mSql As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mAccID      As Integer
        Dim mGAcc       As Integer
        Dim mCnt        As Integer
        mAccID = val(txtAccountHead.Tag)
        mSql = "Select * From faAccountHeads Where intGroupId in(1,2) Order By vchAccountHeadCode Asc"
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF)
                    If mAccID = Rec!intAccountHeadID Then
                        cashBankValidate = True
                        Exit Function
                    End If
                    If txtAccountHeadCode.Text = Rec!vchAccountHeadCode Then
                        cashBankValidate = True
                        Exit Function
                    End If
                    If vsGrid.FindRow(Rec!vchAccountHeadCode, 1, 1, 1, 1) > 0 Then
                         cashBankValidate = True
                         Exit Function
                    End If
                    Rec.MoveNext
                Wend
            End If
            
        End If
    End Function


 
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property

