VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListOfPayments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmListOfPayments"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   15795
      TabIndex        =   22
      Top             =   8310
      Width           =   15855
      Begin VB.CommandButton cmdPayOrder 
         Caption         =   "&View PayOrder"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   15615
      Begin VB.TextBox txtPayOrderNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13080
         TabIndex        =   28
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   12
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "..."
         Height          =   285
         Left            =   7935
         TabIndex        =   9
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox txtInstrumentType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4575
         TabIndex        =   8
         Top             =   360
         Width           =   3345
      End
      Begin VB.TextBox txtBank 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4575
         TabIndex        =   7
         Top             =   690
         Width           =   3345
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11295
         TabIndex        =   6
         Top             =   690
         Width           =   2895
      End
      Begin VB.CommandButton cmdInstrumentType 
         Caption         =   "..."
         Height          =   285
         Left            =   7935
         TabIndex        =   5
         Top             =   360
         Width           =   285
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "..."
         Height          =   270
         Left            =   14220
         TabIndex        =   4
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox txtVrNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11295
         TabIndex        =   3
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtAmount1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11295
         TabIndex        =   2
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtAmount2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12885
         TabIndex        =   1
         Top             =   1020
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60620801
         CurrentDate     =   40197
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   2220
         TabIndex        =   30
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60620801
         CurrentDate     =   40197
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12540
         TabIndex        =   29
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblPayOrderGeneratedSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   20
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   19
         Top             =   375
         Width           =   915
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   300
         TabIndex        =   18
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblForwardedSeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank/Treasury"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   3180
         TabIndex        =   17
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblPaymentOrderNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10275
         TabIndex        =   16
         Top             =   330
         Width           =   975
      End
      Begin VB.Label lblTransactionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9750
         TabIndex        =   15
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10575
         TabIndex        =   14
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12660
         TabIndex        =   13
         Top             =   825
         Width           =   225
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6045
      Left            =   90
      TabIndex        =   21
      Top             =   390
      Width           =   15675
      _cx             =   27649
      _cy             =   10663
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfPayments.frx":0000
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
   Begin VSFlex8LCtl.VSFlexGrid vsPO 
      Height          =   7935
      Left            =   60
      TabIndex        =   32
      Top             =   30
      Visible         =   0   'False
      Width           =   15675
      _cx             =   27649
      _cy             =   13996
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfPayments.frx":0139
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
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "...List of Payments...."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   15750
   End
End
Attribute VB_Name = "frmListOfPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmdBank_Click()
        Dim mSql As String
        
        mSql = " Select A.intAccountHeadID,A.vchBankName +'('+vchAccountHeadCode +')' From("
        mSql = mSql + " Select 1504 as intAccountHeadID,'Cash ' as vchBankName"
        mSql = mSql + " Union All Select intAccountHeadID,vchBankName From faBanks) A"
        mSql = mSql + " Inner Join faAccountHeads On faAccountHeads.intAccountHeadID=A.intAccountHeadID"
        
        frmSearchMasters.SQLQry = mSql
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtBank.Text = gbSearchStr
            txtBank.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdInstrumentType_Click()
        frmSearchMasters.SQLQry = "Select intInstrumentTypeID, vchInstrumentType From faInstrumentTypes"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtInstrumentType.Text = gbSearchStr
            txtInstrumentType.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdNew_Click()
        frmIntegratedPayments.mWebExtract = False
        frmIntegratedPayments.Visible = True
    End Sub

Private Sub cmdPayOrder_Click()
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mArrIn      As Variant
        Dim mSql        As String
        Dim mFromDate   As String
        Dim mToDate     As String
        Dim mStatus     As Variant
        Dim mFwdSeat     As Variant
        vsPO.Visible = True
        fraSearch.Visible = False
        vsGrid.Visible = False
        vsPO.Rows = 1
            mSql = " Select *,faPayOrder.vchDescription As Descriptions,faPayOrder.tnyStatus As Status,faPayOrder.numSeatID As SeatID  From faPayOrder" & vbNewLine
            mSql = mSql + " Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID And faPayOrderChild.tnyCategoryFlag = 3" & vbNewLine
            mSql = mSql + " Inner Join faTransactionType ON faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID" & vbNewLine
            mSql = mSql + " Inner Join faUser On faUser.numUserID = faPayOrder.numUserID" & vbNewLine
            mSql = mSql + " Left Join faVouchers On faVouchers.intVoucherID = faPayOrder.intVoucherID" & vbNewLine
            mSql = mSql + " Where (tnyCancelled <> 1 Or tnyCancelled Is Null)" & vbNewLine
           ' mSql = mSql + " And isnull(numFwdSeatID,0)  Like    '" & IIf(txtForwardedSeat.Tag = "", "%", txtForwardedSeat.Tag) & "'" & vbNewLine
           ' mSql = mSql + " And faPayOrder.numSeatID   LIke '" & IIf(txtGeneratedSeat.Tag = "", "%", txtGeneratedSeat.Tag) & "'" & vbNewLine
            If txtDateFrom.Text <> "" And txtDateTo.Text <> "" Then
                mSql = mSql + " And dtPayOrderDate between '" & DdMmmYy(txtDateFrom.Text) & "' And '" & DdMmmYy(txtDateTo.Text) & "'"
            End If
            'mSql = mSql + " And dtPayOrderDate Between '" & mFromDate & "' And '" & mToDate & "'" & vbNewLine
            'mSql = mSql + " And faPayOrder.tnyStatus in ( " & mStatus & ")" & vbNewLine
            'mSql = mSql + " And faPayOrder. intTransactionTypeID Like '" & IIf(txtTransactionType.Tag = "", "%", txtTransactionType.Tag) & "'" & vbNewLine
            'mSql = mSql + " And faPayOrder. vchPayOrderNo Like '" & IIf(txtPayOrderNo.Text = "", "%", txtPayOrderNo.Text) & "'" & vbNewLine
            'mSQL = mSQL + " And numAmount  Between '" & IIf(txtAmount1.Text = "", "%", val(txtAmount1.Text)) & "' And  '" & val(txtAmount2.Text) & "'" & vbNewLine
            mSql = mSql + " Order By faPayOrder.vchPayOrderNo Desc"
            Set Rec = objdb.ExecuteSP(mSql, , , False, mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                vsPO.Rows = vsPO.Rows + 1
                vsPO.TextMatrix(vsPO.Rows - 1, 0) = IIf(IsNull(DdMmmYy(Rec!dtPayOrderDate)), "", DdMmmYy(Rec!dtPayOrderDate))
                vsPO.TextMatrix(vsPO.Rows - 1, 1) = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                vsPO.TextMatrix(vsPO.Rows - 1, 2) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsPO.TextMatrix(vsPO.Rows - 1, 3) = IIf(IsNull(Rec!Descriptions), "", Rec!Descriptions) ' Description
                vsPO.TextMatrix(vsPO.Rows - 1, 4) = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
                vsPO.TextMatrix(vsPO.Rows - 1, 5) = IIf(IsNull(Rec!SeatID), "", GetSeatName(Rec!SeatID))
                vsPO.TextMatrix(vsPO.Rows - 1, 6) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                If Rec!Status = 1 Then
                    vsPO.Cell(flexcpChecked, vsPO.Rows - 1, 7) = vbChecked
                Else
                    vsPO.Cell(flexcpChecked, vsPO.Rows - 1, 7) = vbUnchecked
                End If
                vsPO.TextMatrix(vsPO.Rows - 1, 8) = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                vsPO.TextMatrix(vsPO.Rows - 1, 9) = IIf(IsNull(Rec!intModuleID), "", Rec!intModuleID)
                vsPO.TextMatrix(vsPO.Rows - 1, 10) = IIf(IsNull(Rec!numFwdSeatID), "", GetSeatName(Rec!numFwdSeatID))
                If Rec!Status = 1 Or Rec!Status = 3 Then
                    vsPO.Cell(flexcpChecked, vsPO.Rows - 1, 11) = vbChecked
                Else
                    vsPO.Cell(flexcpChecked, vsPO.Rows - 1, 11) = vbUnchecked
                End If
                Rec.MoveNext
            Wend
        End If
End Sub
 Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        On Error GoTo err:
        objSeatName.CreateNewConnection mCnnSeatName, enuSourceString.DBMaster
        
        mSQLSeatName = "Select * From GL_Seats"
        mSQLSeatName = mSQLSeatName + " Where numSeatID = " & mSeatID
        RecSeatName.Open mSQLSeatName, mCnnSeatName
        If Not (RecSeatName.EOF And RecSeatName.BOF) Then
            GetSeatName = IIf(IsNull(RecSeatName!chvSeatTitle), "", RecSeatName!chvSeatTitle)
        End If
        RecSeatName.Close
        Exit Function
err:
        MsgBox (Error$)
    End Function



    Private Sub cmdsearch_Click()
        fraSearch.Visible = True
        vsGrid.Visible = True
        vsPO.Visible = False
        FillPayment
    End Sub

    Private Sub cmdTransactionType_Click()
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        frmSearchTransactionType.ModeOfTransaction = 2
        frmSearchTransactionType.Show vbModal
        If gbSearchID > 0 Then
            txtTransactionType.Text = Trim(gbSearchStr)
            txtTransactionType.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    Private Sub cmdView_Click()
       
        If vsGrid.Row > 0 Then
            Dim aryIn As Variant
            aryIn = Array(vsGrid.TextMatrix(vsGrid.Row, 8))
            frmViewVoucher.ArrayIn = aryIn
            frmViewVoucher.FormName = "PaymentVoucher"
            frmViewVoucher.Show vbModal
         End If
    End Sub
    Private Sub dtpDateFrom_CloseUp()
        txtDateFrom.Text = dtpDateFrom.Value
        txtDateFrom.SetFocus
    End Sub
    
    Private Sub dtpDateTo_CloseUp()
        txtDateTo.Text = dtpDateTo.Value
        txtDateTo.SetFocus
    End Sub

    Private Sub Form_Activate()
        dtpDateFrom = gbStartingDate
        dtpDateTo = gbTransactionDate
        
         'Me.Left = 0
        'Me.Top = 0
        Me.WindowState = 2
        '-----------------------------------------------------'
        '                   Form Load Code                    '
        '-----------------------------------------------------'
        
        Me.WindowState = 2
        txtDateFrom.Text = Date - 31
        If CDate(txtDateFrom.Text) < gbStartingDate Then
            txtDateFrom.Text = gbStartingDate
        End If
        txtDateTo.Text = Date
        txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
        txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)

        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdNew.Visible = True
        Else
            cmdNew.Visible = False
        End If
        FillPayment
    End Sub

    
    Private Sub FillPayment()
        Dim mSql   As String
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mTrType As Variant
        Dim mDate   As String
        Dim mProject As Variant
        Dim mAmt    As Variant

        mSql = " Select dtDate,intVoucherNo,fltAmount,vchAccountHead +'('+ vchAccountHeadCode+')' vchAccountHead,vchInstrumentNo,intKeyID2,vchTransactionType,isNull(tnyReversed,0) tnyReversed,intVoucherID"
        mSql = mSql + " From faVouchers"
        mSql = mSql + " Inner Join faAccountHeads On faAccountHeads.intAccountheadID=faVouchers.intKeyID1"
        mSql = mSql + " Inner Join faTransactionType On faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID"
        mSql = mSql + " Where tnyVoucherTypeID = 20 And isnUll(tnyStatus,0)=0"
        If txtDateFrom.Text <> "" And txtDateTo.Text <> "" Then
                mSql = mSql + " And dtDate between '" & DdMmmYy(txtDateFrom.Text) & "' And '" & DdMmmYy(txtDateTo.Text) & "'"
        End If
        If txtInstrumentType.Tag <> "" Then
            mSql = mSql + " And intInstrumentTypeID = " & val(txtInstrumentType.Tag)
        End If
        If txtBank.Tag <> "" Then
            mSql = mSql + " And intKeyID1 = " & val(txtBank.Tag)
        End If
        If txtVrNo.Text <> "" Then
            mSql = mSql + " And intVoucherNo = " & val(txtVrNo.Text)
        End If
        If txtPayOrderNo.Text <> "" Then
            mSql = mSql + " And intKeyID2 = " & val(txtPayOrderNo.Text)
        End If
        If txtTransactionType.Tag <> "" Then
            mSql = mSql + " And faVouchers.intTransactionTypeID = " & val(txtTransactionType.Tag)
        End If
  
        If txtAmount1.Text <> "" And txtAmount2.Text <> "" Then
                mSql = mSql + " And fltAmount between " & txtAmount1.Text & " And " & txtAmount2.Text
        End If
        
        mSql = mSql + " Order By dtDate Desc,intVoucherNo Desc"
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(DdMmmYy(Rec!dtDate)), "", DdMmmYy(Rec!dtDate))
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead) ' Description
                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                'vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!tnyReversed), "", Rec!tnyReversed)
                If Rec!tnyReversed = 1 Then
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 7) = vbChecked
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, 7) = &H9696FF
                Else
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 7) = vbUnchecked
                End If
                vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
         
                Rec.MoveNext
            Wend
        End If
        Rec.Close

    End Sub
    Private Sub txtAmount1_KeyPress(KeyAscii As Integer)
         If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
                KeyAscii = 0
            End If
    End Sub
    
    Private Sub txtAmount2_KeyPress(KeyAscii As Integer)
         If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
                KeyAscii = 0
            End If
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            frmIntegratedPayments.DisplayVoucherDetails (vsGrid.TextMatrix(vsGrid.Row, 1))
        End If
    End Sub
