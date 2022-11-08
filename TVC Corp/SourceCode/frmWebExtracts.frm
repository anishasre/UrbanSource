VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWebExtracts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WebExtracts"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Height          =   1860
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   15915
      Begin VB.CheckBox chkEbill 
         Caption         =   "E-Bills without Voucher"
         Height          =   285
         Left            =   4860
         TabIndex        =   29
         Top             =   1170
         Width           =   2475
      End
      Begin VB.CommandButton cmdVoucher 
         Caption         =   "View Voucher"
         Height          =   375
         Left            =   12450
         TabIndex        =   28
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton cmdViewVoucher 
         Caption         =   "View"
         Height          =   375
         Left            =   14400
         TabIndex        =   26
         Top             =   330
         Width           =   1485
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   255
         Left            =   3900
         TabIndex        =   25
         Top             =   1530
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Report"
         Height          =   1515
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   4485
         Begin VB.ComboBox cmbCategory 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   900
            Width           =   3555
         End
         Begin VB.ComboBox cmbSource 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   510
            Width           =   3555
         End
         Begin VB.ComboBox cmbYear 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   150
            Width           =   2325
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            Caption         =   "Source"
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
            Left            =   210
            TabIndex        =   24
            Top             =   480
            Width           =   585
         End
         Begin VB.Label lblCategory 
            AutoSize        =   -1  'True
            Caption         =   "Category"
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
            TabIndex        =   23
            Top             =   930
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Year"
            Height          =   210
            Left            =   330
            TabIndex        =   22
            Top             =   180
            Width           =   420
         End
      End
      Begin VB.CheckBox chkP 
         Alignment       =   1  'Right Justify
         Caption         =   "P"
         Height          =   195
         Left            =   4590
         TabIndex        =   11
         Top             =   780
         Width           =   465
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "R"
         Height          =   195
         Left            =   4590
         TabIndex        =   1
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6180
         TabIndex        =   4
         Top             =   750
         Width           =   2805
      End
      Begin VB.CommandButton cmdlinkebillwithRP 
         Caption         =   "LinkEbill"
         Height          =   375
         Left            =   9060
         TabIndex        =   10
         Top             =   1170
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   7380
         TabIndex        =   9
         Top             =   1170
         Width           =   1545
      End
      Begin VB.CommandButton CancelList 
         Caption         =   "Cancel List"
         Height          =   375
         Left            =   14370
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   14370
         TabIndex        =   15
         Top             =   1050
         Width           =   1545
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10200
         TabIndex        =   5
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10200
         TabIndex        =   6
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtebill 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6180
         TabIndex        =   3
         Top             =   450
         Width           =   2805
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   330
         Left            =   11400
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60751873
         CurrentDate     =   42929
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   330
         Left            =   11400
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60751873
         CurrentDate     =   42929
      End
      Begin VB.Label Label1 
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
         Left            =   5250
         TabIndex        =   17
         Top             =   750
         Width           =   675
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
         Left            =   9240
         TabIndex        =   14
         Top             =   345
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
         Left            =   9480
         TabIndex        =   13
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblPaymentOrderNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E Bill No"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   450
         Width           =   825
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5955
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   15855
      _cx             =   27966
      _cy             =   10504
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
      BackColorAlternate=   14408667
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmWebExtracts.frx":0000
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
   Begin VB.Label lblPreyear 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Year Mode"
      Height          =   315
      Left            =   5970
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "frmWebExtracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' faWebExtracts intWebExtractType 0=Normal,1=Vr generated,2=VrLinkRequest,3=LinkVerify,4=LinkApprove
Public mPreviousYearMode As Integer
Public mWebExtractDate As Date
Public mWebSourceID As Integer
Public mSubLedgerID As Double
Function AutoWordWrap(vs As VSFlexGrid)
        With vs
            If .Rows > 1 Then
                .AutoSizeMode = flexAutoSizeRowHeight
                .WordWrap = True
                .AutoSize 0, .Cols - 1
                .Cell(flexcpAlignment, 1, 5, .Rows - 1, .Cols - 1) = 0
            End If
        End With
    End Function
    Public Sub FillLinkDetails()
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecSub      As New ADODB.Recordset
        Dim mRowCnt     As Integer
        
     
        If objdb.SetConnection(mCnn) Then
            If mPreviousYearMode = 1 Then
                mSql = " SELECT *,faWebExtracts.tnyVoucherTypeID,faWebExtracts.fltAmount Amt,faWebExtracts.dtDate as webDate,"
                mSql = mSql + " faVouchers.dtDate as VrDate ,faWebExtracts.intTransactionTypeID as TrTypeID From faWebExtracts Left Join faVouchers on faWebExtracts.numKeyID=faVouchers.intVoucherId"
                mSql = mSql + " Where faWebExtracts.intFinancialyearId=" & gbFinancialYearID - 1
            Else
                mSql = " SELECT *,faWebExtracts.tnyVoucherTypeID,faWebExtracts.fltAmount Amt,faWebExtracts.dtDate as webDate,"
                mSql = mSql + " faVouchers.dtDate as VrDate ,faWebExtracts.intTransactionTypeID as TrTypeID From faWebExtracts Left Join faVouchers on faWebExtracts.numKeyID=faVouchers.intVoucherId"
                mSql = mSql + " Where faWebExtracts.intFinancialyearId=" & gbFinancialYearID
            End If
                mRowCnt = 1
            
                If txtebill.Text <> "" Then
                    mSql = mSql + "  And faWebExtracts.numbillcontrolcode= '" & txtebill.Text & "'"
                End If
                If txtAmount.Text <> "" Then
                    mSql = mSql + "  And faWebExtracts.fltAmount= " & txtAmount.Text
                End If
                If txtDateFrom.Text <> "" Then
                    If txtDateTo.Text <> "" Then
                        mSql = mSql + "  And CONVERT(smalldatetime,CONVERT(char(11), faWebExtracts.dtDate))  between '" & DdMmmYy(txtDateFrom.Text) & "' And '" & DdMmmYy(txtDateTo.Text) & "'"
                    Else
                        MsgBox "please Enter To Date", vbApplicationModal
                        Exit Sub
                    End If
                End If
                If chkP.Value = 1 Then
                    mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=2 "
                End If
                If chkR.Value = 1 Then
                    mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=1 "
                End If
                mSql = mSql + " Order By faWebExtracts.dtDate,numbillcontrolcode,numbillcontrolID,faWebExtracts.tnyVoucherTypeID Desc "
                 Rec.CursorLocation = adUseClient
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                vsGrid.Clear 1, 1
                vsGrid.Rows = 1
                
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    If IsNull(Rec!webDate) Then
                        vsGrid.TextMatrix(mRowCnt, 0) = "Synchronised Date Error"
                    Else
                        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!webDate), "", Format(Rec!webDate, "DD-MMM-YYYY"))
                    
                    End If
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!Amt), "", Rec!Amt)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
                    If (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 1 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "R"
                        
                        
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 2 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "P"
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 4 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "JV"
                    End If
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VrDate), "", Format(Rec!VrDate, "DD-MMM-YYYY"))
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
                    vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intExtractTypeID), 0, Rec!intExtractTypeID)
                    If IsNull(Rec!webDate) Then
                        vsGrid.TextMatrix(mRowCnt, 11) = 100
                    End If
                    If val(vsGrid.TextMatrix(mRowCnt, 11)) = 0 Then
                        vsGrid.TextMatrix(mRowCnt, 10) = "."
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Voucher generated"
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 2 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Requested for Voucher Link"
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 3 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Link Request Verified "
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 4 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Link with Vr Done"
                    End If
                    vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numbillcontrolID), 0, Rec!numbillcontrolID)
                    vsGrid.MergeCells = flexMergeFixedOnly
                    vsGrid.MergeCol(12) = True
                   
                    vsGrid.MergeRow(mRowCnt) = True
                    vsGrid.WordWrap = True
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                Wend
                Call AutoWordWrap(vsGrid)
                Rec.Close
        End If
    End Sub
    Private Sub FillGrid()
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecSub      As New ADODB.Recordset
        Dim mRowCnt     As Integer
        
     
        If objdb.SetConnection(mCnn) Then
                If mPreviousYearMode = 1 Then
                    mSql = " SELECT *,faWebExtracts.tnyVoucherTypeID,faWebExtracts.fltAmount Amt,faWebExtracts.dtDate as webDate,"
                    mSql = mSql + " faVouchers.dtDate as VrDate ,faWebExtracts.intTransactionTypeID as TrTypeID From faWebExtracts Left Join faVouchers on faWebExtracts.numKeyID=faVouchers.intVoucherId"
                    mSql = mSql + " Where faWebExtracts.intFinancialyearId=" & gbFinancialYearID - 1
                Else
                    mSql = " SELECT *,faWebExtracts.tnyVoucherTypeID,faWebExtracts.fltAmount Amt,faWebExtracts.dtDate as webDate,"
                    mSql = mSql + " faVouchers.dtDate as VrDate ,faWebExtracts.intTransactionTypeID as TrTypeID From faWebExtracts Left Join faVouchers on faWebExtracts.numKeyID=faVouchers.intVoucherId"
                    mSql = mSql + " Where faWebExtracts.intFinancialyearId=" & gbFinancialYearID
                End If
                mRowCnt = 1
            
                If txtebill.Text <> "" Then
                    mSql = mSql + "  And faWebExtracts.numbillcontrolcode= '" & txtebill.Text & "'"
                End If
                If txtAmount.Text <> "" Then
                    mSql = mSql + "  And faWebExtracts.fltAmount= " & txtAmount.Text
                End If
                If txtDateFrom.Text <> "" Then
                    If txtDateTo.Text <> "" Then
                        mSql = mSql + "  And CONVERT(smalldatetime,CONVERT(char(11), faWebExtracts.dtDate))  between '" & DdMmmYy(txtDateFrom.Text) & "' And '" & DdMmmYy(txtDateTo.Text) & "'"
                    Else
                        MsgBox "please Enter To Date", vbApplicationModal
                        Exit Sub
                    End If
                End If
                If chkP.Value = 1 Then
                    mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=2 "
                End If
                If chkR.Value = 1 Then
                    mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=1 "
                End If
                '' Added on 3 /may/2019
                If chkEbill.Value = 1 Then
                    mSql = mSql + "  And faWebExtracts.numKeyID is null "
                End If
                mSql = mSql + " Order By faWebExtracts.dtDate,numbillcontrolcode,numbillcontrolID,faWebExtracts.tnyVoucherTypeID Desc "
                 Rec.CursorLocation = adUseClient
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                vsGrid.Clear 1, 1
                vsGrid.Rows = 1
                
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    If IsNull(Rec!webDate) Then
                        vsGrid.TextMatrix(mRowCnt, 0) = "Synchronised Date Error"
                    Else
                        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!webDate), "", Format(Rec!webDate, "DD-MMM-YYYY"))
                    
                    End If
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!Amt), "", Rec!Amt)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
                    If (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 1 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "R"
                        
                        
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 2 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "P"
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 4 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "JV"
                    End If
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VrDate), "", Format(Rec!VrDate, "DD-MMM-YYYY"))
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
                    vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intExtractTypeID), 0, Rec!intExtractTypeID)
                    If IsNull(Rec!webDate) Then
                        vsGrid.TextMatrix(mRowCnt, 11) = 100
                    End If
                    If val(vsGrid.TextMatrix(mRowCnt, 11)) = 0 Then
                        vsGrid.TextMatrix(mRowCnt, 10) = "."
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Voucher generated"
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 2 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Requested for Voucher Link"
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 3 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Link Request Verified "
                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 4 Then
                         vsGrid.TextMatrix(mRowCnt, 10) = "Link with Vr Done"
                    End If
                    vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numbillcontrolID), 0, Rec!numbillcontrolID)
                    vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intImplofficerTypeID), 0, Rec!intImplofficerTypeID)
                    vsGrid.MergeCells = flexMergeFixedOnly
                    vsGrid.MergeCol(12) = True
                   
                    vsGrid.MergeRow(mRowCnt) = True
                    vsGrid.WordWrap = True
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                Wend
                Call AutoWordWrap(vsGrid)
                Rec.Close
                
       End If
    End Sub

    Private Sub CancelList_Click()
        frmCancelListofWebExtractVoucher.Show vbModal
    End Sub




    Private Sub chkEbill_Click()
        If chkEbill.Value = 1 Then
             chkR.Value = 0
        End If
    End Sub

    Private Sub chkP_Click()
        If chkP.Value = 1 Then
            chkR.Value = 0
        End If
    End Sub

    Private Sub chkR_Click()
        If chkR.Value = 1 Then
            chkP.Value = 0
        End If
    End Sub



Private Sub cmdCancel_Click()
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    Dim Rec1 As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    If objdb.SetConnection(mCnn) Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 6)) > 1 Then
            
                mSql = "Select faWebExtractChild.intAccountHeadID,faVoucherChild.intAccountHeadID,* From faWebExtracts"
                mSql = mSql + " Inner Join faWebExtractChild On faWebExtractChild.intWebExtractID=faWebExtracts.intWebExtractID"
                mSql = mSql + " Inner Join faVouchers On faWebExtracts.numKeyID=faVouchers.intVoucherID"
                mSql = mSql + " Inner Join faVoucherChild On faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                mSql = mSql + " Where faWebExtracts.tnyVoucherTypeID in (1, 2) And faWebExtractChild.intSlNo <> 1"
                mSql = mSql + " And faWebExtractChild.intAccountHeadID<>faVoucherChild.intAccountHeadID and faWebExtracts.numKeyID is not Null"
                mSql = mSql + " and numBillControlCode in ( '" & vsGrid.TextMatrix(vsGrid.Row, 1) & "')"
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                If Not (Rec.EOF And Rec.BOF) Then
                    If (MsgBox("Account head not matching . Do you want to Cancel", vbOKCancel)) = vbOK Then
                        mSql = " Update faVouchers SET tnyCancelFlag = 1, tnyStatus = 4 WHERE intVoucherID IN (Select numKeyId From faWebExtracts Where  numBillControlCode=" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                        mSql = mSql + " Update faTransactions SET tnyStatus = 4,tnyReversed=Null WHERE intVoucherID IN (Select numKeyId From faWebExtracts Where  numBillControlCode=" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                        mSql = mSql + " Update faWebExtracts set numKeyId=Null,intExtractTypeID=Null Where numBillControlCode in(" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                        mCnn.Execute mSql
                        
                    Else
                    
                    End If
                    MsgBox "Cancelled Successfully", vbApplicationModal
                    Call FillGrid
                    Exit Sub
                    
                Else
                    
                    mSql = " Select * From faWebExtracts "
                    mSql = mSql + " Inner Join faVouchers On faWebExtracts.numKeyID=faVouchers.intVoucherID"
                    mSql = mSql + " Where faWebExtracts.fltAmount <> faVouchers.fltAmount"
                    mSql = mSql + " and numBillControlCode in ( '" & vsGrid.TextMatrix(vsGrid.Row, 1) & "')"
                    Rec1.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                    If Not (Rec1.EOF And Rec1.BOF) Then
                        If (MsgBox("Amount is not matching . Do you want to Cancel", vbOKCancel)) = vbOK Then
                            mSql = " Update faVouchers SET tnyCancelFlag = 1, tnyStatus = 4 WHERE intVoucherID IN (Select numKeyId From faWebExtracts Where  numBillControlCode=" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                            mSql = mSql + " Update faTransactions SET tnyStatus = 4,tnyReversed=Null WHERE intVoucherID IN (Select numKeyId From faWebExtracts Where  numBillControlCode=" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                            mSql = mSql + " Update faWebExtracts set numKeyId=Null,intExtractTypeID=Null Where numBillControlCode in(" & vsGrid.TextMatrix(vsGrid.Row, 1) & ")"
                            mCnn.Execute mSql
                            
                        End If
                        MsgBox "Cancelled Successfully", vbApplicationModal
                        Call FillGrid
                        Exit Sub
                    Else
                        MsgBox "Selected E bill's Amount or Acc head Are same with generated Voucher ,So its not allowed to cancel", vbInformation
                        Exit Sub
                    End If
                End If
                            
    '            frmCancelWebExtractVoucher.DispayWebExtractVoucher (vsGrid.TextMatrix(vsGrid.Row, 8))
    '            frmCancelWebExtractVoucher.Show
            Else
                MsgBox "Please Select Voucher Generated E - bill", vbInformation
            End If
        End If
    End If
End Sub

Private Sub cmdlinkebillwithRP_Click()
    'If gbSeatGroupID = gbSeatGroupAccountsClerk Then
        If vsGrid.Row < 1 Then
            MsgBox "please Select an e bill", vbApplicationModal
            Exit Sub
        End If
        If val(vsGrid.TextMatrix(vsGrid.Row, 6)) < 1 Then
            If mPreviousYearMode = 1 Then
                frmLinkEbillWithRP.mPreYearMode = 1
            End If
            frmLinkEbillWithRP.cmdLinkRP.Tag = val(vsGrid.TextMatrix(vsGrid.Row, 11))
            frmLinkEbillWithRP.DispayWebExtractTolinkVoucher (vsGrid.TextMatrix(vsGrid.Row, 8))
            frmLinkEbillWithRP.Show vbModal
        Else
            MsgBox "Select E - bill which has no Voucher Number", vbInformation
        End If
    'End If
End Sub

Private Sub cmdsearch_Click()

Call FillGrid
'    Dim objAcc As New clsAccounts
'    Dim mSql As String
'    Dim mcnn As New ADODB.Connection
'    Dim Rec As New ADODB.Recordset
'    Dim RecChild As New ADODB.Recordset
'    Dim objDb As New clsDB
'    Dim mCount As Integer
''    frmIntegratedPayments.mWebExtract = True
'
''    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
''            MsgBox "Connction not Present ", vbCritical
''            Exit Sub
''    End If
'
'    mSql = " SELECT *,faWebExtracts.tnyVoucherTypeID,faWebExtracts.fltAmount Amt,faWebExtracts.dtDate as webDate,faVouchers.dtDate as VrDate "
'    mSql = mSql + " From faWebExtracts Left Join faVouchers on faWebExtracts.numKeyID=faVouchers.intVoucherId  "
'    mSql = mSql + " Where numbillcontrolcode is not null"
'    If txtebill.Text <> "" Then
'        mSql = mSql + "  And faWebExtracts.numbillcontrolcode= '" & txtebill.Text & "'"
'    End If
'    If txtAmount.Text <> "" Then
'        mSql = mSql + "  And faWebExtracts.fltAmount= " & txtAmount.Text
'    End If
'    If txtDateFrom.Text <> "" Then
'        If txtDateTo.Text <> "" Then
'            mSql = mSql + "  And CONVERT(smalldatetime,CONVERT(char(11), faWebExtracts.dtDate))  between '" & DdMmmYy(txtDateFrom.Text) & "' And '" & DdMmmYy(txtDateTo.Text) & "'"
'        Else
'            MsgBox "please Enter To Date", vbApplicationModal
'            Exit Sub
'        End If
'    End If
'    If chkP.value = 1 Then
'        mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=2 "
'    End If
'    If chkR.value = 1 Then
'        mSql = mSql + "  And faWebExtracts.tnyVoucherTypeID=1 "
'    End If
'    mSql = mSql + " Order By faWebExtracts.dtDate,numbillcontrolcode "
'    If objDb.SetConnection(mcnn) Then
'     Rec.CursorLocation = adUseClient
'                Rec.Open mSql, mcnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
'                mRowCnt = 1
'
'                vsGrid.Clear 1, 1
'                vsGrid.Rows = 1
'                While Not (Rec.EOF Or Rec.BOF)
'                    vsGrid.Rows = vsGrid.Rows + 1
'                    If IsNull(Rec!webDate) Then
'                         vsGrid.TextMatrix(mRowCnt, 0) = "Synchronised Date Error"
'                    Else
'                        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!webDate), "", CheckDateInMMM(Rec!webDate))
'                    End If
'                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
'                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!Amt), "", Rec!Amt)
'                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
'                    If (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 1 Then
'                        vsGrid.TextMatrix(mRowCnt, 4) = "R"
'                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 2 Then
'                        vsGrid.TextMatrix(mRowCnt, 4) = "P"
'                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 4 Then
'                        vsGrid.TextMatrix(mRowCnt, 4) = "JV"
'                    End If
'                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
'                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VrDate), "", Format(Rec!VrDate, "DD-MMM-YYYY"))
'                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
'                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
'                    vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intExtractTypeID), 0, Rec!intExtractTypeID)
'                    If IsNull(Rec!webDate) Then
'                        vsGrid.TextMatrix(mRowCnt, 11) = 100
'                    End If
'                    If val(vsGrid.TextMatrix(mRowCnt, 11)) = 0 Then
'                        vsGrid.TextMatrix(mRowCnt, 10) = "."
'                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 1 Then
'                         vsGrid.TextMatrix(mRowCnt, 10) = "Voucher generated"
'                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 2 Then
'                         vsGrid.TextMatrix(mRowCnt, 10) = "Requested for Voucher Link"
'                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 3 Then
'                         vsGrid.TextMatrix(mRowCnt, 10) = "Link Request Verified "
'                    ElseIf vsGrid.TextMatrix(mRowCnt, 11) = 4 Then
'                         vsGrid.TextMatrix(mRowCnt, 10) = "Link with Vr Done"
'                    End If
'                    vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numbillcontrolID), 0, Rec!numbillcontrolID)
'                    Rec.MoveNext
'                    mRowCnt = mRowCnt + 1
'                Wend
'                Rec.Close
'  Call AutoWordWrap(vsGrid)
'    End If
End Sub



Private Sub cmdView_Click()
      Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCategory As String
        Dim mYearID As Integer
        Dim mSource As String
        If cmbYear.ListIndex > -1 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
        Else
            mYearID = gbFinancialYearID
        End If
        
        If cmbCategory.ListIndex < 1 Then
        mCategory = "%"
            
        Else
            mCategory = cmbCategory.ItemData(cmbCategory.ListIndex)
        End If
        If cmbSource.ListIndex < 1 Then
'            MsgBox "Please select the Source", vbInformation
'            cmbSource.SetFocus
'            Exit Sub
'            arInput = Array(mCategory, cmbSource.ItemData(cmbSource.ListIndex), mYearID)
            mSource = "%"
        Else
            mSource = cmbSource.ItemData(cmbSource.ListIndex)
     
        End If
            arInput = Array(mYearID, mSource, mCategory)
            frmNewViewer.rptFileName = App.Path & "\Reports\rptEbillReport.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show

End Sub

    Private Sub cmdViewVoucher_Click()
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        If vsGrid.Row < 1 Then
            MsgBox "please Select an e bill", vbApplicationModal
            Exit Sub
        End If
        If vsGrid.TextMatrix(vsGrid.Row, 12) > 1 Then
            arInput = Array(vsGrid.TextMatrix(vsGrid.Row, 12))
            frmNewViewer.rptFileName = App.Path & "\Reports\rptEbillVrDetails.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show
        End If
    End Sub

    Private Sub cmdVoucher_Click()
           
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCategory As String
        Dim mYearID As Integer
        Dim mSource As String
        
        If vsGrid.Row < 1 Then
            MsgBox "please Select an e bill", vbApplicationModal
            Exit Sub
        End If
        If vsGrid.TextMatrix(vsGrid.Row, 6) > 1 Then
            arInput = Array(vsGrid.TextMatrix(vsGrid.Row, 6))
            frmNewViewer.rptFileName = App.Path & "\Reports\rptVoucher.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show
        Else
            MsgBox " Voucher not Generated", vbInformation
        End If
    End Sub

    Private Sub dtpDateFrom_CloseUp()
        txtDateFrom.Text = DdMmmYy(dtpDateFrom.Value)
    End Sub

    Private Sub dtpDateTo_CloseUp()
        txtDateTo.Text = DdMmmYy(dtpDateTo.Value)
    End Sub

    Private Sub Form_Activate()
        If mPreviousYearMode = 1 Then
            lblPreyear.Visible = True
        Else
            lblPreyear.Visible = False
        End If
        Call FillGrid
    End Sub

Private Sub FillPayment(mWebId As Double)
    Dim objAcc As New clsAccounts
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mCount As Integer
    Dim onjTrType As New clsTransactionType
    frmIntegratedPayments.mWebExtract = True
    If mPreviousYearMode = 1 Then
        frmIntegratedPayments.mPreYearMode = 1
    End If
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    mSql = " SELECT  *,isnull(faWebExtracts.numKeyID,0) as VrID From faWebExtracts  Inner Join faWebExtractChild On  faWebExtracts.intwebExtractID=faWebExtractChild.intwebExtractID"
    mSql = mSql + " Inner Join faWebExtractAddress On faWebExtractAddress.intwebExtractID=faWebExtracts.intwebExtractID"
    mSql = mSql + " Inner Join suSourceOfFund On suSourceOfFund.intSourceFundID=faWebExtracts.intSourceID"
    mSql = mSql + " Where faWebExtracts.intwebExtractID= " & vsGrid.TextMatrix(vsGrid.Row, 8)

    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        
        frmIntegratedPayments.txtDate = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
        frmIntegratedPayments.txtDated = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
        If mPreviousYearMode = 1 Then
        frmIntegratedPayments.mPreYearMode = 1
    End If
        frmIntegratedPayments.txtInstrument.Text = "Cash"
        frmIntegratedPayments.txtInstrument.Tag = 1
        frmIntegratedPayments.txtInstrument.Enabled = False
        frmIntegratedPayments.txtCrHeadCode.Text = gbAcHeadCodeCash
        frmIntegratedPayments.txtCrHeadCode.Tag = 1504
        frmIntegratedPayments.txtCrAccountHead.Text = "Cash"
        frmIntegratedPayments.cmdCrAccountHead.Enabled = False
        frmIntegratedPayments.cmdInstrument.Enabled = False
        frmIntegratedPayments.txtCrAmount.Text = val(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount))
        frmIntegratedPayments.lblTotal.Caption = val(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount))
        frmIntegratedPayments.txtNarration.Text = "Project:" & IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
        frmIntegratedPayments.txtNarration.Enabled = False
        frmIntegratedPayments.txtBillControCodeID.Text = IIf(IsNull(Rec!numbillcontrolID), "", Rec!numbillcontrolID)
        frmIntegratedPayments.txtWebExtractIDforP.Tag = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
        If gbLBPanchayat = 1 Then
            
            'frmIntegratedPayments.txtFunctionary.Text = "101 Secretary, Village Panchayat"
            'frmIntegratedPayments.txtFunctionary.Tag = 1
            
        Else
            'frmIntegratedPayments.txtFunctionary.Text = "010000 Secretary's Section"
            'frmIntegratedPayments.txtFunctionary.Tag = 1
        End If
        frmIntegratedPayments.txtSourceofFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
        frmIntegratedPayments.txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
        Call FillFunctionary
        onjTrType.SetTransactionType (val(IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)))
        frmIntegratedPayments.txtTransactionType.Text = onjTrType.TransactionType '"Development Project Expenditure -General- Capital"
        frmIntegratedPayments.txtTransactionType.Tag = val(IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID))
        frmIntegratedPayments.cmdSearchTransactionType.Enabled = False
        frmIntegratedPayments.txtName.Text = IIf(IsNull(Rec!numbillcontrolcode), ".", Rec!numbillcontrolcode)
        'frmIntegratedPayments.txtPayOrder.Text = IIf(IsNull(Rec!numbillcontrolcode), ".", Rec!numbillcontrolcode)
        frmIntegratedPayments.cmdPaymentOrder.Enabled = False
        frmIntegratedPayments.txtPayOrder.Enabled = False
        frmIntegratedPayments.txtAgreementNo.Enabled = False
        frmIntegratedPayments.cmdSearchVoucher.Enabled = False
        frmIntegratedPayments.cmdAgreementNo.Enabled = False
        frmIntegratedPayments.cmdAllotmentLetterNo.Enabled = False
        frmIntegratedPayments.cmdImplementingOfficer.Enabled = False
        frmIntegratedPayments.cmdNew.Enabled = False
        frmIntegratedPayments.txtVoucherNo.Enabled = False
        frmIntegratedPayments.txtAccountNo.Text = ""
        frmIntegratedPayments.txtInstrumentNo.Text = ""
        frmIntegratedPayments.txtAccountNo.Enabled = False
        frmIntegratedPayments.txtInstrumentNo.Enabled = False
        frmIntegratedPayments.txtName.Tag = vsGrid.TextMatrix(vsGrid.Row, 8)
        frmIntegratedPayments.txtName.Enabled = False
        frmIntegratedPayments.txtBranch.Enabled = False
        frmIntegratedPayments.txtNameOfBank.Text = ""
        If (Rec!VrID) > 0 Then
            frmIntegratedPayments.cmdSave.Enabled = False
            frmIntegratedPayments.txtVoucherNo = vsGrid.TextMatrix(vsGrid.Row, 5)
        End If
    End If
    Rec.Close
            
        'frmIntegratedPayments.vsGrid.Rows = 1
        frmIntegratedPayments.vsGrid.MergeCells = flexMergeFree
       ' frmIntegratedPayments.vsGrid.Editable = flexEDNone
        mSql = "SELECT  * From faWebExtractChild   Inner Join faAccountHeads On faAccountHeads.intAccountHeadID=faWebExtractChild.intAccountHeadID Where faWebExtractChild.intwebExtractID=" & vsGrid.TextMatrix(vsGrid.Row, 8)
        RecChild.Open mSql, mCnn
          If Not (RecChild.EOF And RecChild.BOF) Then
                mCount = 0
                While Not RecChild.EOF
                    If RecChild!intSlNo <> 1 Then
                 
                        objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                        frmIntegratedPayments.vsGrid.Rows = frmIntegratedPayments.vsGrid.Rows + 1
                        mCount = mCount + 1
                        If objAcc.AccountCode <> gbAcHeadCodeCash Then
                            objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                            frmIntegratedPayments.vsGrid.Cell(flexcpText, mCount, 1) = objAcc.AccountCode
                            frmIntegratedPayments.vsGrid.Cell(flexcpText, mCount, 2) = objAcc.AccountHead
                            frmIntegratedPayments.vsGrid.Cell(flexcpText, mCount, 3) = val(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount))
                            frmIntegratedPayments.vsGrid.Cell(flexcpText, mCount, 4) = objAcc.AccountHeadID
                            frmIntegratedPayments.vsGrid.Row = mCount
                        End If
                    End If
                    RecChild.MoveNext
                  Wend
        End If
        mCnn.Close
End Sub
Private Function FillFunctionary() As String
    Dim SQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Function
        End If
        mSql = "Select faFunctionaries.intFunctionaryID intFun,* From faFunctionaryWebImplOfficer "
        mSql = mSql + " Inner Join faFunctionaries On faFunctionaries.intFunctionaryID=faFunctionaryWebImplOfficer.intFunctionaryID"
        mSql = mSql + " Where intWebImpID = " & vsGrid.TextMatrix(vsGrid.Row, 13)
        'chvImplOfficerDesgEng vchFunctionaryCode
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            frmIntegratedPayments.txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionaryCode), "", Rec!vchFunctionaryCode) + " " + IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            frmIntegratedPayments.txtFunctionary.Tag = IIf(IsNull(Rec!intFun), "", Rec!intFun)
            frmIntegratedPayments.cmdSearchFunctionary.Enabled = False
        Else
            frmIntegratedPayments.cmdSearchFunctionary.Enabled = False
            If gbLBPanchayat = 1 Then
                If gbLBType = 2 Then
                    frmIntegratedPayments.txtFunctionary.Text = "101 Secretary, Block Panchayat"
                Else
                    frmIntegratedPayments.txtFunctionary.Text = "101 Secretary, Village Panchayat"
                End If
                frmIntegratedPayments.txtFunctionary.Tag = 1
  
            Else
                frmIntegratedPayments.txtFunctionary.Text = "010000 Secretary's Section"
                frmIntegratedPayments.txtFunctionary.Tag = 1
            End If
        End If
End Function
Private Sub FillJournal(mWebId As Double)
        Dim mCount As Long
        Dim mFineWaveDate As String
        Dim objAcc As New clsAccounts
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim objdb As New clsDB
        
        frmJournalEntry.mWebExtractJV = True
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
        End If
        
        mSql = " SELECT  *,isnull(faWebExtracts.numKeyID,0) as VrID From faWebExtracts  Inner Join faWebExtractChild On  faWebExtracts.intwebExtractID=faWebExtractChild.intwebExtractID"
        mSql = mSql + " Inner Join faWebExtractAddress On faWebExtractAddress.intwebExtractID=faWebExtracts.intwebExtractID"
        mSql = mSql + " Where faWebExtracts.intwebExtractID= " & vsGrid.TextMatrix(vsGrid.Row, 8)

        Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmJournalEntry.mWebExtractJVDate = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
                frmJournalEntry.txtDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
                'frmJournalEntry.Tag = vsGrid.TextMatrix(vsGrid.Row, 8)
                frmJournalEntry.vsGrid.Editable = flexEDNone
                frmJournalEntry.cmbTransactionType.Text = "Dev Pr Exp -General- Capital"
                frmJournalEntry.cmbTransactionType.Tag = 1141
                'frmJournalEntry.cmb = IIf(IsNull(Rec!numbillcontrolcode), ".", Rec!numbillcontrolcode)
                'frmJournalEntry.txtAccountHeadCode.Text
                frmJournalEntry.txtBudgetCentreCode.Tag = vsGrid.TextMatrix(vsGrid.Row, 8)
                frmJournalEntry.txtNarration = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
                frmJournalEntry.txtNarration.Enabled = False
                If gbLBPanchayat = 1 Then
                    frmJournalEntry.txtFund.Text = "Panchayat Fund"
                Else
                    frmJournalEntry.txtFund.Text = "General Fund"
                End If
                frmJournalEntry.txtFund.Tag = 1
                frmJournalEntry.cmdNew.Enabled = False
                If (Rec!VrID) > 0 Then
                    frmJournalEntry.cmdSave.Enabled = False
                    frmJournalEntry.txtVoucherNo = vsGrid.TextMatrix(vsGrid.Row, 5)
                End If

            End If
            Rec.Close
            
        frmJournalEntry.vsGrid.MergeCells = flexMergeFree
        frmJournalEntry.vsGrid.Editable = flexEDNone
        mSql = "SELECT  * From faWebExtractChild   Inner Join faAccountHeads On faAccountHeads.intAccountHeadID=faWebExtractChild.intAccountHeadID Where faWebExtractChild.intwebExtractID=" & vsGrid.TextMatrix(vsGrid.Row, 8)
        RecChild.Open mSql, mCnn
          If Not (RecChild.EOF And RecChild.BOF) Then
                mCount = 0
                While Not RecChild.EOF
                    If RecChild!intSlNo = 1 Then
                        If IIf(IsNull(RecChild!tnyDebitCreditFlag), "", RecChild!tnyDebitCreditFlag) = 1 Then
                            frmJournalEntry.optDebit.Value = True
                        Else
                            frmJournalEntry.optCredit.Value = True
                        End If
                        objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                        frmJournalEntry.txtAccountHeadCode.Text = objAcc.AccountCode
                        frmJournalEntry.txtAccountHead.Tag = objAcc.AccountHeadID
                        frmJournalEntry.txtAccountHead.Text = objAcc.AccountHead
                        frmJournalEntry.cmdSearchAccountHead.Enabled = False

                    Else
                        mCount = mCount + 1
                        objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                        frmJournalEntry.vsGrid.Cell(flexcpText, mCount, 1) = objAcc.AccountCode
                        frmJournalEntry.vsGrid.Cell(flexcpText, mCount, 2) = objAcc.AccountHead
                        frmJournalEntry.vsGrid.Cell(flexcpText, mCount, 4) = val(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount))

                    End If
                   
                  RecChild.MoveNext
                  Wend
       
            End If
End Sub
Private Sub FillReceipt(mWebId As Double)

        Dim mCount As Long
        Dim mFineWaveDate As String
        Dim objAcc As New clsAccounts
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim objdb As New clsDB
        
        
        frmReceiptsCounter.mWebExtractMode = True
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
        End If

        mSql = " SELECT  *,isnull(faWebExtracts.numKeyID,0) as VrID From faWebExtracts  Inner Join faWebExtractChild On  faWebExtracts.intwebExtractID=faWebExtractChild.intwebExtractID"
        mSql = mSql + " Inner Join faWebExtractAddress On faWebExtractAddress.intwebExtractID=faWebExtracts.intwebExtractID"
        mSql = mSql + " Where faWebExtracts.intwebExtractID= " & vsGrid.TextMatrix(vsGrid.Row, 8)

        Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
            
     
                frmReceiptsCounter.mWebExtractDate = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
                'mWebExtractDate = DdMmmYy(IIf(IsNull(Rec!dtDate), gbTransactionDate, Rec!dtDate))
                frmReceiptsCounter.txtDate = DdMmmYy(mWebExtractDate)
                frmReceiptsCounter.txtOutDoorStaff.Tag = vsGrid.TextMatrix(vsGrid.Row, 8)
                frmReceiptsCounter.vsGrid.Editable = flexEDNone
                mWebSourceID = IIf(IsNull(Rec!intSourceID), -1, Rec!intSourceID)
                frmReceiptsCounter.txtName.Text = IIf(IsNull(Rec!numbillcontrolcode), ".", Rec!numbillcontrolcode)
                frmReceiptsCounter.txtTransactionType.Text = "Development Project Expenditure -General- Capital"
                frmReceiptsCounter.txtTransactionType.Tag = 1141
                frmReceiptsCounter.txtDescription = "Project:" & IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
                frmReceiptsCounter.txtDescription.Enabled = False
                frmReceiptsCounter.txtWardNo.Enabled = False
                frmReceiptsCounter.txtDoorNo1.Enabled = False
                frmReceiptsCounter.txtDoorNo2.Enabled = False
                frmReceiptsCounter.txtName.Enabled = False
                frmReceiptsCounter.txtInit1.Enabled = False
                frmReceiptsCounter.txtInit2.Enabled = False
                frmReceiptsCounter.txtInit3.Enabled = False
                frmReceiptsCounter.txtInit4.Enabled = False
                frmReceiptsCounter.txtHouse.Enabled = False
                frmReceiptsCounter.txtStreet.Enabled = False
                frmReceiptsCounter.txtLocalPlace.Enabled = False
                frmReceiptsCounter.txtMainPlace.Enabled = False
                frmReceiptsCounter.txtPost.Enabled = False
                frmReceiptsCounter.txtPin.Enabled = False
                frmReceiptsCounter.txtPhone.Enabled = False
                
                If (Rec!VrID) > 0 Then
                    frmReceiptsCounter.cmdSave.Enabled = False
                    frmReceiptsCounter.txtReceiptNo = vsGrid.TextMatrix(vsGrid.Row, 5)
                End If

            End If
            Rec.Close
            
        frmReceiptsCounter.vsGrid.Rows = 1
        frmReceiptsCounter.vsGrid.MergeCells = flexMergeFree
        mSql = "SELECT  * From faWebExtractChild   Inner Join faAccountHeads On faAccountHeads.intAccountHeadID=faWebExtractChild.intAccountHeadID Where faWebExtractChild.intwebExtractID=" & vsGrid.TextMatrix(vsGrid.Row, 8)
        RecChild.Open mSql, mCnn
          If Not (RecChild.EOF And RecChild.BOF) Then
                mCount = 0
                While Not RecChild.EOF
                    If RecChild!intSlNo = 1 Then
                        objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                        frmReceiptsCounter.txtAccountHead.Text = objAcc.AccountHead & "[" & objAcc.AccountCode & "]"
                        frmReceiptsCounter.txtAccountHead.Tag = objAcc.AccountHeadID
                        'If objAcc.AccountCode = gbAcHeadCodeCash Then
                            frmReceiptsCounter.txtInstrument.Text = "Cash"
                            frmReceiptsCounter.txtInstrument.Tag = 1
                            frmReceiptsCounter.cmdSearchInstrument.Enabled = False
                            frmReceiptsCounter.cmdSearchAccountHead.Enabled = False
'                            frmReceiptsCounter.txtInstrument.Enabled = False
'                        Else
'                            frmReceiptsCounter.txtInstrument.Enabled = False
'                            frmReceiptsCounter.txtInstrument.Text = ""
'                            frmReceiptsCounter.txtInstrument.Tag = -1
'                        End If
                    Else
                        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
                        mCount = mCount + 1
                        objAcc.SetAccountCode (val(IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)))
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 0) = objAcc.AccountCode
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 1) = objAcc.AccountHead
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 5) = val(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount))
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 6) = objAcc.AccountHeadID
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 12) = 1
                        
                        frmReceiptsCounter.vsGrid.Row = mCount
                    End If
                   
                  RecChild.MoveNext
                  Wend
                  frmReceiptsCounter.Calculate
            End If
            'Unload Me
        End Sub



Private Sub Form_Load()
     Dim mSql As String
        
        If gbLBPanchayat = 1 Then
            mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,25,26,27,28,10,11,12,13,14,19,21,29,30,41,42,43,44,45) Order By vchSourceFundName"
        Else
            mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,19,21,25,26,27,28,29,30,41,42,43,44,45) Order By vchSourceFundName"
        End If
        PopulateList cmbSource, mSql, , True, True, True, enuSourceString.Saankhya
        
        mSql = "SELECT vchTransactionCategory,intCategoryID FROM faTransactionCategory"
        PopulateList cmbCategory, mSql, True, True, True, True
        
      
        On Error Resume Next
        mSql = "Select LTRIM(Str(intFinancialYear)) + '-' + LTRIM(Str(intFinancialYear+1)), intFinancialYearID  From faFinancialYear Where intFinancialYear>2016"
        PopulateList cmbYear, mSql, True, False, True, True
        cmbYear.Text = Trim(str(gbFinancialYearID)) + "-" + Trim(str((gbFinancialYearID + 1)))
        
        Call FillGrid
End Sub


Private Sub txtAmount_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
End Sub

Private Sub txtDateFrom_LostFocus()
    txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
End Sub

Private Sub txtDateTo_LostFocus()
    txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
End Sub

Private Sub txtebill_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
End Sub
    Private Sub vsGrid_DblClick()
    If vsGrid.Row > 0 Then
        If vsGrid.TextMatrix(vsGrid.Row, 11) = 0 Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 1 Then  'Type 1 For Receipt 2 for Payment 4 for journal.
                'If gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                '    Call FillReceipt(vsGrid.TextMatrix(vsGrid.Row, 8))
'                Else
'                    MsgBox "Cachier/Chief Casheir can do Receipt"
'                    Exit Sub
'                End If
            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 2 Then
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
      
                    Call FillPayment(vsGrid.TextMatrix(vsGrid.Row, 8))
                Else
                    MsgBox "AccountsClerk/Accountant can do Payment"
                    Exit Sub
                End If
'            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 4 Then
'                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
'
'                    Call FillJournal(vsGrid.TextMatrix(vsGrid.Row, 8))
'                Else
'                    MsgBox "AccountsClerk/Accountant can do Payment"
'                    Exit Sub
'                End If
            End If
        ElseIf vsGrid.TextMatrix(vsGrid.Row, 11) = 100 Then
            MsgBox "Inncorrect Synchronised data"
            Exit Sub
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 1 Then
            MsgBox "Voucher Already Generated", vbApplicationModal
            Exit Sub
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 2 Then
            MsgBox "E bill is Requested to Link with Voucher", vbApplicationModal
            Exit Sub
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 3 Then
            MsgBox "Link Request Verified", vbApplicationModal
            Exit Sub
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 3 Then
            MsgBox "E bill Linked with Voucher", vbApplicationModal
            Exit Sub
        End If
        End If
    End Sub
