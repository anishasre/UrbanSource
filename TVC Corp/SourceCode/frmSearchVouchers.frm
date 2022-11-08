VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchVouchers 
   BackColor       =   &H00FBFFFB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Vouchers"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInterrupted 
      BackColor       =   &H00FBFFFB&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11670
      TabIndex        =   28
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtVoucherNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      TabIndex        =   21
      Top             =   4800
      Width           =   1650
   End
   Begin VB.TextBox txtAmount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      TabIndex        =   20
      Top             =   6525
      Width           =   1650
   End
   Begin VB.TextBox txtInstrumentNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      TabIndex        =   19
      Top             =   5520
      Width           =   1650
   End
   Begin VB.TextBox txtBank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      TabIndex        =   18
      Top             =   5850
      Width           =   5550
   End
   Begin VB.TextBox txtTransactionType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6195
      Width           =   5550
   End
   Begin VB.ComboBox cmbInstrumentType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5145
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FBFFFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12090
      TabIndex        =   12
      Top             =   6870
      Width           =   12150
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   14
         Top             =   30
         Width           =   1470
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         TabIndex        =   13
         Top             =   30
         Width           =   1470
      End
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   10950
         Top             =   345
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   15
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdSearchTransactionType 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   10
      Top             =   6210
      Width           =   285
   End
   Begin VB.TextBox txtAmountTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4260
      TabIndex        =   9
      Top             =   6525
      Width           =   1650
   End
   Begin VB.CheckBox chkJournal 
      BackColor       =   &H00FBFFFB&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11220
      TabIndex        =   8
      Top             =   405
      Width           =   435
   End
   Begin VB.CheckBox chkContra 
      BackColor       =   &H00FBFFFB&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10665
      TabIndex        =   7
      Top             =   405
      Width           =   435
   End
   Begin VB.CheckBox chkPayment 
      BackColor       =   &H00FBFFFB&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10125
      TabIndex        =   6
      Top             =   405
      Width           =   435
   End
   Begin VB.CheckBox chkReceipt 
      BackColor       =   &H00FBFFFB&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9570
      TabIndex        =   5
      Top             =   405
      Width           =   435
   End
   Begin VB.TextBox txtToDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2625
      TabIndex        =   4
      Top             =   375
      Width           =   1140
   End
   Begin VB.TextBox txtFromDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1035
      TabIndex        =   2
      Top             =   375
      Width           =   1140
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3780
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   12090
      _cx             =   21325
      _cy             =   6667
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchVouchers.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instrument No"
      Height          =   195
      Left            =   1020
      TabIndex        =   27
      Top             =   5565
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   195
      Left            =   1815
      TabIndex        =   26
      Top             =   5880
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
      Height          =   195
      Left            =   780
      TabIndex        =   25
      Top             =   6195
      Width           =   1470
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Instrument Type"
      Height          =   195
      Left            =   840
      TabIndex        =   24
      Top             =   5145
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1395
      TabIndex        =   23
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   1590
      TabIndex        =   22
      Top             =   6570
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   4020
      TabIndex        =   11
      Top             =   6570
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   2340
      TabIndex        =   3
      Top             =   420
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   525
      TabIndex        =   0
      Top             =   420
      Width           =   405
   End
End
Attribute VB_Name = "frmSearchVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private intCheckMode As Integer '-----10=Receipt ; 20=Payment ; 30=Contra ; 40=Journal
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearRequestID As Variant
    Public mEbillLinkMode As Boolean
    Private Sub FillGrid()
        Dim mSQL             As String
        Dim objDB            As New clsDB
        Dim mCnn             As New ADODB.Connection
        Dim Rec              As New ADODB.Recordset
        Dim mRow             As Double
        Dim mWhere           As String
        Dim mBankDrawnFrom   As String
        Dim mFlag As Boolean
        Dim mCount As Integer

        mFlag = False
        mCount = 0
        
        If chkReceipt.value = 0 And chkPayment.value = 0 And chkContra.value = 0 And chkJournal.value = 0 And chkInterrupted.value = 0 Then
            MsgBox "Please Select a Voucher Type!", vbInformation
            Exit Sub
        End If
        vsGrid.Rows = 1
        lblCount.Caption = "Rec: 0"
label:
        mSQL = " SELECT CASE  WHEN tnyVoucherGroupID=4  Then cast(intVoucherNo as varchar)+isnull('-'+vchDoorNoP3,'')else cast(intVoucherNo as varchar) END[intVoucherNo], dtDate, vchInstrumentNo, "
        mSQL = mSQL + " vchTransactionType, vchBank, fltAmount,intVoucherID,vchInstrumentType ,vchSectionName"
        mSQL = mSQL + " FROM faVouchers LEFT JOIN  faTransactionType "
        mSQL = mSQL + " ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID  "
        mSQL = mSQL + " LEFT JOIN faInstrumentTypes "
        mSQL = mSQL + " ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID "
        mSQL = mSQL + " LEFT JOIN faCounters ON faCounters.intCounterID=faVouchers.intCounterID"
        mSQL = mSQL + " LEFT JOIN faSection ON faSection.intSectionID=faCounters.intSectionID"
        mWhere = " Where ISNULL(faVouchers.intTransactionTypeID,0) Not in (0,3000) And IsNull(tnyCancelFlag,0) =0 And tnyVoucherTypeID IN ("
        If chkReceipt.value = 1 Then
            mWhere = mWhere + "10,"
        End If
        If chkPayment.value = 1 Then
            mWhere = mWhere + "20,"
        End If
        If chkContra.value = 1 Then
            mWhere = mWhere + "30,"
        End If
        If chkJournal.value = 1 Then
            mWhere = mWhere + "40,"
        End If
        If chkInterrupted.value = 1 Then
            mWhere = mWhere + "10) And tnyVoucherGroupID = 4 "
        Else
            mWhere = Left(mWhere, Len(mWhere) - 1) & ")"
        End If
        If mFlag = False Then
            If (txtFromDate.Text) <> "" And (txtToDate.Text) <> "" Then
                mWhere = mWhere + "And dtDate Between '" & Trim(txtFromDate.Text) & "' And '" & Trim(txtToDate.Text) & "'"
'                mWhere = mWhere + "And dtDate Between '" & Trim(txtFromDate.Text) & "' And convert(varchar(11),cast('" & Trim(txtToDate.Text) & "' as DateTime),103)"
            End If
        Else
            'txtFromDate.Text = DdMmmYy(gbStartingDate)
            'txtToDate.Text = DdMmmYy(gbTransactionDate)
            mWhere = mWhere + "And dtDate Between '" & DdMmmYy(gbStartingDate) & "' And '" & DdMmmYy(gbTransactionDate) & "'"

        End If
        If Trim(txtInstrumentNo.Text) <> "" Then
            mWhere = mWhere + "And faVouchers.vchInstrumentNo LIKE '%" & Trim(txtInstrumentNo.Text) & "%'"
        End If
        
        If val(txtTransactionType.Tag) > 0 Then
'            If mEbillLinkMode = True Then
'                mWhere = mWhere + " And faVouchers.intTransactionTypeID in (1141,1151,1161,1171,1181,1191) "
'            Else
                mWhere = mWhere + " And faVouchers.intTransactionTypeID = " & val(txtTransactionType.Tag)
'            End If
        Else
            If mEbillLinkMode = True Then
                mWhere = mWhere + " And faVouchers.intTransactionTypeID in (1141,1151,1161,1171,1181,1191) "
            End If
        End If
        If Trim(txtBank) <> "" Then
            mBankDrawnFrom = GetDrawnFromBank
            mWhere = mWhere + " And faVouchers.vchBank LIKE '%" & mBankDrawnFrom & "%'"
        End If
        
        If val(txtAmount.Text) <> 0 And val(txtAmountTo.Text) <> 0 Then
            mWhere = mWhere + " And faVouchers.fltAmount BETWEEN " & val(txtAmount.Text) & " And " & val(txtAmountTo.Text)
        ElseIf val(txtAmount.Text) > 0 Then
            mWhere = mWhere + " And faVouchers.fltAmount = " & val(txtAmount.Text)
        ElseIf val(txtAmountTo.Text) > 0 Then
            mWhere = mWhere + " And faVouchers.fltAmount = " & val(txtAmountTo.Text)
        End If
    
        If val(cmbInstrumentType.Tag) > 0 Then
            mWhere = mWhere + " And faVouchers.intInstrumentTypeID = " & val(cmbInstrumentType.Tag)
        End If
        If Trim(txtVoucherNo.Text) <> "" Then
            mWhere = mWhere + " And faVouchers.intVoucherNo LIKE '%" & Trim(txtVoucherNo.Text) & "%'"
'            If vsGrid.TextMatrix(1, 1) <> "" Then
'                txtFromDate.Text = vsGrid.TextMatrix(1, 1)
'            End If
        End If
        
        mSQL = mSQL + mWhere
        mSQL = mSQL + " Order By dtDate"
        
        objDB.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        If Not (Rec.EOF Or Rec.BOF) Then
            vsGrid.MousePointer = flexHourglass
            vsGrid.Rows = Rec.RecordCount + 1
            lblCount.Caption = "Rec:" & str(Rec.RecordCount)
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = 8
            vsGrid.RowSel = vsGrid.Rows - 1
            
            mSQL = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = mSQL
            vsGrid.Row = 1
            vsGrid.Col = 0
            cmdSearch.Enabled = False
            If Rec.EOF Then
                vsGrid.MousePointer = flexDefault
                cmdSearch.Enabled = True
            End If
        Else
            mCount = mCount + 1
            If mCount < 2 Then
                mFlag = True
                mSQL = ""
                Rec.Close
                GoTo label
                
            End If
        End If
        Rec.Close
    End Sub
    
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        
        If mPreviousYearMode = 1 Then
            txtFromDate.Text = DdMmmYy(DateSerial(gbFinancialYearID, 3, 1))
            txtToDate.Text = DdMmmYy(DateSerial(gbFinancialYearID, 3, 31))
        Else
            txtFromDate.Text = DdMmmYy(DateAdd("d", -30, gbTransactionDate))
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        End If
        chkReceipt.value = 0
        chkPayment.value = 0
        chkContra.value = 0
        chkJournal.value = 0
        vsGrid.Clear 1, 0
        gbSearchCode = ""
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Function GetDrawnFromBank()  ' To search Bank Name separated by %
        Dim mLetter As String
        Dim mWord   As String
        Dim mLength As Integer
        Dim mStart  As Integer
        
        mWord = ""
        mLength = Len(txtBank.Text)
        For mStart = 1 To mLength
            If mID(txtBank.Text, mStart, 1) <> " " Then
                mWord = mWord + mID(txtBank.Text, mStart, 1) + "%"
            End If
        Next
        GetDrawnFromBank = mWord
    End Function
    


    Private Sub cmbInstrumentType_Click()          '27/11/2009
        If cmbInstrumentType.ListIndex > -1 Then 'Note:- If any item is selected then
            cmbInstrumentType.Tag = cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
            'Note:-Check whether it is Instrument-Cash
            If cmbInstrumentType.Tag = gbInstrumentCash Then
                'Note:-Disable Instrument No and Name of Bank Text Boxes because of Cash type Instrument
                txtInstrumentNo.Enabled = False
                txtBank.Enabled = False
                txtInstrumentNo.Text = ""
                txtBank.Text = ""
            Else
                txtInstrumentNo.Enabled = True
                txtBank.Enabled = True
            End If
        End If
    End Sub
    
    Private Sub cmdClear_Click()
        vsGrid.Clear 1, 1
        cmbInstrumentType.ListIndex = -1
        txtInstrumentNo.Text = ""
        txtBank.Text = ""
        txtTransactionType.Text = ""
        txtAmount.Text = ""
        txtAmountTo.Text = ""
        lblCount.Caption = ""
        txtVoucherNo.Text = ""
    End Sub
    
    Private Sub cmdsearch_Click()
    Dim mWhere As String
        Call FillGrid
        If Trim(txtVoucherNo.Text) <> "" Then
            mWhere = mWhere + " And faVouchers.intVoucherNo LIKE '%" & Trim(txtVoucherNo.Text) & "%'"
            If vsGrid.Rows > 1 Then
                txtFromDate.Text = DdMmmYy(vsGrid.TextMatrix(1, 1))
            End If
        End If
    
    '    Dim mSql             As String
    '    Dim objDb            As New clsDB
    '    Dim mCnn             As New ADODB.Connection
    '    Dim Rec              As New ADODB.Recordset
    '    Dim mRow             As Double
    '    Dim mWhere           As String
    '    Dim mBankDrawnFrom   As String
    '
    '    If chkReceipt.Value = 0 And chkPayment.Value = 0 And chkContra.Value = 0 And chkJournal.Value = 0 Then
    '        MsgBox "Please Select a Voucher Type!", vbInformation
    '        Exit Sub
    '    End If
    '
    '    vsGrid.Rows = 1
    '    lblCount.Caption = "Rec: 0"
    '
    '    mSql = " SELECT intVoucherNo, dtDate, vchInstrumentNo,vchInstrumentType, "
    '    mSql = mSql + " vchTransactionType, vchBank, fltAmount,intVoucherID "
    '    mSql = mSql + " FROM faVouchers LEFT JOIN  faTransactionType "
    '    mSql = mSql + " ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID  "
    '    mSql = mSql + " INNER JOIN faInstrumentTypes "
    '    mSql = mSql + " ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID "
    '
    '    mWhere = " Where IsNull(tnyCancelFlag,0) =0 And tnyVoucherTypeID IN ("
    '    If chkReceipt.Value = 1 Then
    '        mWhere = mWhere + "10,"
    '    End If
    '    If chkPayment.Value = 1 Then
    '        mWhere = mWhere + "20,"
    '    End If
    '    If chkContra.Value = 1 Then
    '        mWhere = mWhere + "30,"
    '    End If
    '    If chkJournal.Value = 1 Then
    '        mWhere = mWhere + "40,"
    '    End If
    '    mWhere = Left(mWhere, Len(mWhere) - 1) & ")"
    '
    '    If (txtFromDate.Text) <> "" And (txtToDate.Text) <> "" Then
    '        mWhere = mWhere + "And dtDate Between '" & Trim(txtFromDate.Text) & "' And '" & Trim(txtToDate.Text) & "'"
    '    End If
    '    If Trim(txtInstrumentNo.Text) <> "" Then
    '        mWhere = mWhere + "And faVouchers.vchInstrumentNo LIKE '%" & Trim(txtInstrumentNo.Text) & "%'"
    '    End If
    '
    '    If Val(txtTransactionType.Tag) > 0 Then
    '        mWhere = mWhere + " And faVouchers.intTransactionTypeID = " & Val(txtTransactionType.Tag)
    '    End If
    '    If Trim(txtBank) <> "" Then
    '        mBankDrawnFrom = GetDrawnFromBank
    '        mWhere = mWhere + " And faVouchers.vchBank LIKE '%" & mBankDrawnFrom & "%'"
    '    End If
    '
    '    If Val(txtAmount.Text) <> 0 And Val(txtAmountTo.Text) <> 0 Then
    '        mWhere = mWhere + " And faVouchers.fltAmount BETWEEN " & Val(txtAmount.Text) & " And " & Val(txtAmountTo.Text)
    '    ElseIf Val(txtAmount.Text) > 0 Then
    '        mWhere = mWhere + " And faVouchers.fltAmount = " & Val(txtAmount.Text)
    '    End If
    '
    '    If Val(cmbInstrumentType.Tag) > 0 Then
    '        mWhere = mWhere + " And faVouchers.intInstrumentTypeID = " & Val(cmbInstrumentType.Tag)
    '    End If
    '
    '    mSql = mSql + mWhere
    '    mSql = mSql + " Order By dtDate"
    '
    '    objDb.SetConnection mCnn
    '    Rec.CursorLocation = adUseClient
    '    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    '    If Not (Rec.EOF Or Rec.BOF) Then
    '        vsGrid.Rows = Rec.RecordCount + 1
    '        lblCount.Caption = "Rec:" & str(Rec.RecordCount)
    '        vsGrid.Col = 0
    '        vsGrid.Row = 1
    '        vsGrid.ColSel = 6
    '        vsGrid.RowSel = vsGrid.Rows - 1
    '
    '        mSql = Rec.GetString(, , vbTab, Chr(13))
    '        vsGrid.Clip = mSql
    '        vsGrid.Row = 1
    '        vsGrid.Col = 0
    '    End If
    '    Rec.Close
    End Sub
    
    Private Sub cmdSearchTransactionType_Click()
        If (chkReceipt.value = True) Then
           frmSearchTransactionType.ModeOfTransaction = 1
        End If
        If (chkPayment.value = True) Then
           frmSearchTransactionType.ModeOfTransaction = 2
        End If
        frmSearchTransactionType.Show vbModal
        
        txtTransactionType.Text = Trim(gbSearchStr)
        txtTransactionType.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            Call FillGrid
        End If
    End Sub
    
    Private Sub Form_Load()
        XPC.InitIDESubClassing
        Call FillInstrumentType
        FormInitialize
        
        'To use this searchVoucher from Another Form
        If (CheckMode = 10) Then
            chkReceipt.value = vbChecked
        ElseIf CheckMode = 20 Then
            chkPayment = vbChecked
        ElseIf CheckMode = 30 Then
            chkContra = vbChecked
        ElseIf CheckMode = 40 Then
            chkJournal = vbChecked
        End If
        
        
        
    End Sub
    Private Sub FillInstrumentType()
        Dim mSqlIns As String
        mSqlIns = "SELECT vchInstrumentType,intInstrumentTypeID from faInstrumentTypes"
        PopulateList cmbInstrumentType, mSqlIns, , True, True, True
    End Sub
    
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If

    End Sub

    Private Sub txtAmount_LostFocus()
        If val(txtAmount) > 0 Then
            txtAmount.Text = Format(val(txtAmount.Text), "0.00")
        Else
            txtAmount.Text = ""
        End If
    End Sub
    
    Private Sub txtAmountTo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtAmountTo_LostFocus()
        If val(txtAmountTo) > 0 Then
            txtAmountTo.Text = Format(val(txtAmountTo), "0.00")
        Else
            txtAmountTo.Text = ""
        End If
    End Sub
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        If IsDate(txtFromDate) Then
            If mPreviousYearMode = 1 Then
                If Not (CDate(txtToDate.Text) >= CDate(DateAdd("yyyy", -1, gbStartingDate)) And CDate(txtToDate.Text) <= CDate(DateAdd("yyyy", -1, gbEndingDate))) Then
                    txtToDate.Text = DdMmmYy(DateSerial(gbFinancialYearID, 3, 31))
                End If
                If Not (CDate(txtFromDate.Text) >= CDate(DateAdd("yyyy", -1, gbStartingDate)) And CDate(txtFromDate.Text) <= CDate(DateAdd("yyyy", -1, gbEndingDate))) Then
                    txtFromDate.Text = DdMmmYy(DateSerial(gbFinancialYearID - 1, 4, 1))
                End If
            Else
                If txtFromDate < gbStartingDate Or txtFromDate > gbEndingDate Then
                    txtFromDate.Text = DdMmmYy(gbStartingDate)
                End If
            End If
        End If
        
    End Sub
    
    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        If IsDate(txtToDate.Text) Then
            If mPreviousYearMode = 1 Then
                If Not (CDate(txtToDate.Text) >= CDate(DateAdd("yyyy", -1, gbStartingDate)) And CDate(txtToDate.Text) <= CDate(DateAdd("yyyy", -1, gbEndingDate))) Then
                    txtToDate.Text = DdMmmYy(DateSerial(gbFinancialYearID, 3, 31))
                End If
                If Not (CDate(txtFromDate.Text) >= CDate(DateAdd("yyyy", -1, gbStartingDate)) And CDate(txtFromDate.Text) <= CDate(DateAdd("yyyy", -1, gbEndingDate))) Then
                    txtFromDate.Text = DdMmmYy(DateSerial(gbFinancialYearID - 1, 4, 1))
                End If
            Else
                If txtToDate.Text < gbStartingDate Or txtToDate.Text > gbEndingDate Then
                    txtToDate.Text = DdMmmYy(gbStartingDate)
                End If
            End If
        End If
    End Sub
    
    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer) 'To delete value in txtTransactionType
        If KeyCode = vbKeyDelete Then
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        End If
    End Sub
    Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsGrid_DblClick()
     If (vsGrid.TextMatrix(vsGrid.Row, 0) <> "") Then
        gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 0) 'Voucher No
        gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 3)  'Transaction Type
        gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 6)   'Voucher ID
        gbReceiptDate = vsGrid.TextMatrix(vsGrid.Row, 1) 'Voucher Date
        Unload Me
     End If
    End Sub
    
    Public Property Let CheckMode(mData As Integer)
        intCheckMode = mData
    End Property
    Public Property Get CheckMode() As Integer
        CheckMode = intCheckMode
    End Property
    

    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property

