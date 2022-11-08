VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmZonalDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmZonalDetails"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeContra 
      Caption         =   "Contra Entry Posting"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   30
      TabIndex        =   14
      Top             =   8070
      Width           =   14295
      Begin VB.TextBox txtAccountCode 
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
         Height          =   375
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   570
         Width           =   1695
      End
      Begin VB.TextBox txtAccountHead 
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
         Height          =   375
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   500
         TabIndex        =   21
         Top             =   570
         Width           =   5595
      End
      Begin VB.CommandButton cmdAccoundHeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "..."
         Height          =   375
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txtVoucherNo 
         Enabled         =   0   'False
         Height          =   345
         Left            =   9690
         TabIndex        =   17
         Top             =   570
         Width           =   1845
      End
      Begin VB.TextBox txtContraAmount 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8010
         TabIndex        =   16
         Top             =   540
         Width           =   1335
      End
      Begin VB.CommandButton cmdContra 
         Caption         =   "Save "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         TabIndex        =   15
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher Number"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Cr. A/c Head"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8010
         TabIndex        =   18
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.Frame frmS 
      Height          =   1095
      Left            =   45
      TabIndex        =   5
      Top             =   6945
      Width           =   14295
      Begin VB.CheckBox chkWrongVouchers 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "VERIFICATION  FAILED  LIST"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   255
         TabIndex        =   13
         Top             =   735
         Width           =   2460
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdTransafer 
         Caption         =   "Transtfer"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtTotalCash 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtTotalBankAmt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   14175
      Begin VSFlex8LCtl.VSFlexGrid VSGridZonalDetails 
         Height          =   5775
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   14115
         _cx             =   24897
         _cy             =   10186
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
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
         Rows            =   2
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalDetails.frx":0000
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
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         Height          =   345
         Left            =   12375
         TabIndex        =   3
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox txtZonalName 
         Enabled         =   0   'False
         Height          =   345
         Left            =   11010
         TabIndex        =   1
         Top             =   345
         Width           =   345
      End
      Begin VB.Label lblZonal 
         Caption         =   "SREEKARIYAM ZONAL OFFICE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   12
         Top             =   285
         Width           =   10725
      End
      Begin VB.Label lblDate 
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11865
         TabIndex        =   2
         Top             =   405
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmZonalDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Dim mnumZonDetailsID, mvarZonName, mdtZonDate, mZonTotal As Variant
    Dim mLoop As Integer
    Dim mYearID As Integer
    
    Private Sub chkWrongVouchers_Click()
        Call FillGrid
    End Sub

    Private Sub cmdAccoundHeads_Click()
        ShowSearchAccountHead
    End Sub
    Private Sub ShowSearchAccountHead()
            Dim mSql As String
            
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID =" & faBank
            frmSearchAccountHeads.VoucherMode = 300
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.chkListAll.Enabled = False
            frmSearchAccountHeads.cmdSearch.Enabled = False
            frmSearchAccountHeads.Show vbModal
            txtAccountCode.SetFocus
            
            txtAccountCode.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1

    End Sub

    Private Sub cmdClose_Click()
        Unload Me
        frmZonalMain.Show
    End Sub
    
    Private Function GetYearID(mDt As Date) As Integer
    
        Dim mYearID As Integer
        Dim mMonthID As Integer
        mYearID = Year(mDt)
        mMonthID = Month(mDt)
        If mMonthID < 4 Then
            mYearID = mYearID - 1
        End If
        GetYearID = mYearID
    
    End Function
    Private Sub FillGrid()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim mRowCount As Variant
        Dim aryIn(1) As Variant
        
        aryIn(0) = mdtZonDate
        aryIn(1) = mnumZonDetailsID
        objdb.CreateNewConnection mCnn, SaankhyaHO
        Set Rec = objdb.ExecuteSP("spGetVouchersData", aryIn, , , mCnn)
        If Month(txtDate) > 3 Then
            mYearID = Year(txtDate)
        Else
            mYearID = Year(txtDate) - 1
        End If
        
        mRowCount = 1
        VSGridZonalDetails.Clear 1, 1
        VSGridZonalDetails.Rows = 1
        While Not Rec.EOF
            VSGridZonalDetails.Rows = VSGridZonalDetails.Rows + 1
            VSGridZonalDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            VSGridZonalDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            VSGridZonalDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            VSGridZonalDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            VSGridZonalDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!cashAmt), "", Rec!cashAmt)
            VSGridZonalDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!BankAmt), "", Rec!BankAmt)
            VSGridZonalDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo) + " " + IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            VSGridZonalDetails.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            VSGridZonalDetails.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!intVoucherIDNew), "", Rec!intVoucherIDNew)
            VSGridZonalDetails.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!tnyCancelFlag), 0, Rec!tnyCancelFlag)
            VSGridZonalDetails.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!intTransactionTypeID), 0, Rec!intTransactionTypeID)
            If Not IsNull(Rec!tnySyncStatus) Then
                If Rec!tnyVerified = 1 Then
                    VSGridZonalDetails.TextMatrix(mRowCount, 7) = vbChecked
                ElseIf Rec!tnyVerified = 0 Then
                    VSGridZonalDetails.TextMatrix(mRowCount, 7) = vbUnchecked
                    VSGridZonalDetails.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 11) = &H8080FF
                    cmdVerify.Enabled = True
                    cmdTransafer.Enabled = False
                Else
                    VSGridZonalDetails.TextMatrix(mRowCount, 7) = vbUnchecked
                    cmdVerify.Enabled = True
                    cmdTransafer.Enabled = False
                End If
                '' Added By anisha On 18.Sep.2014
                If VSGridZonalDetails.TextMatrix(mRowCount, 12) = 1 Then
                    VSGridZonalDetails.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 12) = vbRed
                End If
                If (Rec!tnySyncStatus = 2) Then
                    VSGridZonalDetails.TextMatrix(mRowCount, 8) = vbChecked
                    VSGridZonalDetails.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 11) = &HC0FFC0
                Else
                    If cmdVerify.Enabled = False Then
                        cmdTransafer.Enabled = True
                    End If
                End If
            End If
            If chkWrongVouchers.Value = 1 Then
                If Rec!tnyVerified = 1 Then
                    VSGridZonalDetails.RowHidden(mRowCount) = True
                End If
            End If
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
        Rec.Close
       
        Call Calculation
        mCnn.Close
        Call GetContraAmt
    End Sub

Private Sub cmdContra_Click()

Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        Dim Rec                 As New ADODB.Recordset
        Dim mCnn                As ADODB.Connection
        Dim objdb               As New clsDB
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mSql                As String
    If val(txtAccountHead.Tag) < 1 Then
        MsgBox "Please select Bank for the Zonal collection", vbApplicationModal
        Exit Sub
    End If

     arrInput = Array( _
                    IIf(txtVoucherNo.Tag = "", -1, txtVoucherNo.Tag), _
                    gbLocalBodyID, _
                    Null, _
                    4001, _
                    30, _
                    Null, _
                    Null, _
                    Format(txtDate.Text, "DD/MmM/YYYY"), _
                    val(txtContraAmount.Text), _
                    1, _
                    "Zonal Collection Of " & mvarZonName, _
                    Null, _
                    "Zonal Collection Of " & mvarZonName, _
                    mnumZonDetailsID, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    gbUserID, _
                    gbCounterID, _
                    Null, _
                    1504, Null, 115, _
                    60, _
                    mYearID, Null, Null, Null, Null, Null, 1, gbSeatID, gbSessionID, "Zonal Collection", Null, Null, Null, Null, mnumZonDetailsID, Null, Null)
            
                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID = arrOutPut(0, 0)
                    If arrOutPut(0, 0) <> "" Then
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
                
                mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & mintVoucherID
              
                        
                        mintLocalBodyID_2 = gbLocalBodyID
                        mintSlNo_3 = 1
                        mintAccountHeadID_4 = val(txtAccountHead.Tag)
                        mintYearID_6 = mYearID
                        mtnyPeriodID_7 = 0
                        mtnyArrearFlag_8 = 0
                        mnumDemandID_9 = Null
                        mfltAmount_10 = val(txtContraAmount.Text)
                        
                        arrInput = Array( _
                        mintVoucherID, _
                        mintLocalBodyID_2, _
                        mintSlNo_3, _
                        mintAccountHeadID_4, _
                        1, _
                        mintYearID_6, _
                        mtnyPeriodID_7, _
                        mtnyArrearFlag_8, _
                        mnumDemandID_9, _
                        mfltAmount_10 _
                        )
                        objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
            
                    
        
        
                '-------------------------------------------------------'
                ' faTransactions
                '-------------------------------------------------------'
                arrInput = Array(-1, _
                           gbLocalBodyID, _
                           mYearID, _
                           Format(txtDate.Text, "DD/MmM/YYYY"), _
                           0, _
                           60, _
                           6, _
                           4, _
                           Null, _
                           1, _
                           Null, _
                           "Zonal Collection Of " & mvarZonName, _
                           Null, _
                           0, _
                           "C", _
                           30, _
                           Null, _
                           Null, _
                           gbUserID, _
                           mintVoucherID _
                           )
                
                Rec.CursorLocation = adUseClient
                Call objdb.ExecuteSP("spSaveTransactions", arrInput, arrOutPut, , mCnn)
                
                '-------------------------------------------------------'
                ' faTransactionChild
                '-------------------------------------------------------'
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintTransactionID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack:
                End If
               
                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mintTransactionID

                arrInput = Array(mintTransactionID, _
                            1, _
                            1504, _
                            Format(val(txtContraAmount.Text), "0.00"), _
                            0, _
                            Null, _
                            "Zonal Collection Of " & mvarZonName, _
                            1 _
                            )
                     objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn


                     arrInput = Array(mintTransactionID, _
                             2, _
                             val(txtAccountHead.Tag), _
                             Format(val(txtContraAmount.Text), "0.00"), _
                             1, _
                             1504, _
                             Null, _
                             1 _
                             )
                     objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                     
                     MsgBox "Contra Entry Saved Successfully", vbApplicationModal
                     cmdContra.Enabled = False
                     Exit Sub
ErrRollBack:
        txtVoucherNo.Text = ""
''        cmdSave.Enabled = False
        Debug.Print Error$
        mCnn.RollbackTrans
    
End Sub
    Private Sub GetContraAmt()
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New Recordset
        Dim mRowCount As Variant
        Dim mAmount As Double
        
        
        
      
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = "Select *  From faVouchers "
                mSql = mSql + "Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
                mSql = mSql + " Inner Join faAccountHeads On faAccountHeads.intAccountHeadID=faVoucherChild.intAccountHeadID"
                mSql = mSql + " Where faVouchers.intFinancialYearID=" & gbFinancialYearID & " and tnyVoucherTypeID=30 And intTransactionTypeID=4001 And numLocationId=" & mnumZonDetailsID & " and dtDate='" & txtDate.Text & " '"
'                aryIn(0) = mnumZonDetailsID
'            aryIn(1) = gbFinancialYearID
'            aryIn(2) = mdtZonDate
       ' objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'Set Rec = objdb.ExecuteSP("spgetzonalcontra", aryIn, , , mCnn)
                
               Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    mAmount = Rec!fltAmount
                    txtVoucherNo.Text = Rec!intVoucherNo
                    txtVoucherNo.Tag = Rec!intVoucherID
                    txtAccountHead.Text = Rec!vchAccountHead
                    txtAccountHead.Tag = Rec!intAccountHeadID
                    txtAccountCode.Text = Rec!vchAccountHeadCode
                    cmdContra.Enabled = False
                Else
                        cmdContra.Enabled = True
                End If
                
                Rec.Close
        End If
            mCnn.Close
        
    End Sub
    
    Private Sub VerifyContraAmt()
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New Recordset
        Dim mRowCount As Variant
        Dim mZoneAmountinFinance As Double
        Dim mZoneAmountinHO As Double
        
        
        
      
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select Sum(fltAmount) ZoneAmt  From faVouchers Where tnyVoucherTypeID=10 And numLocationId=" & mnumZonDetailsID & " and dtDate='" & txtDate.Text & " '"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                mZoneAmountinFinance = IIf(IsNull(Rec!ZoneAmt), 0, Rec!ZoneAmt)
            
            End If
            
            Rec.Close
        End If
        mCnn.Close
        If objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO) Then
            mSql = "Select Sum(fltAmount) ZoneAmt  From faVouchers Where tnyVoucherTypeID=10 And numLocationId=" & mnumZonDetailsID & " and dtDate='" & txtDate.Text & " '"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                mZoneAmountinHO = IIf(IsNull(Rec!ZoneAmt), 0, Rec!ZoneAmt)
            End If
            Rec.Close
        End If
        mCnn.Close
        
        If mZoneAmountinFinance <> mZoneAmountinHO Then
            MsgBox "Transfered Zonal amount and Amount in Zonal office (HO) are not Equal"
            cmdContra.Enabled = False
            Exit Sub
        Else
            cmdContra.Enabled = True
        End If
        
    End Sub
    Private Sub cmdTransafer_Click()
        Dim mVoucher            As uVoucher
        Dim mVouChildTbl        As uVChild
        Dim mVouAddress         As uVoucherAddress
        Dim mVoucherSub         As uVoucherSub
        
        Dim mTranTable          As uTr
        Dim mTranChildTbl       As uTrChild
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mCnnHO As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim mRowCount As Variant
        Dim aryIn, aryOut, aryIn1 As Variant
        Dim mLoop, mLoopHO, mCount As Integer
        Dim RLoop, mVoucherID, mTransactionID As Variant
        
        For mLoop = 1 To VSGridZonalDetails.Rows - 1

            aryIn = ""
            aryIn1 = ""
            aryOut = ""
            If val(VSGridZonalDetails.TextMatrix(mLoop, 8)) <> vbChecked Then
                If val(VSGridZonalDetails.TextMatrix(mLoop, 7)) = vbChecked Then
                    '------------------Vouchers------------
                
                    objdb.CreateNewConnection mCnnHO, enuSourceString.SaankhyaHO
                    mSql = ""
                    mSql = "SELECT * FROM faVouchers WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND numLocationID=" & ZonalDetailsID
                    Rec.Open mSql, mCnnHO
                    If objdb.SetConnection(mCnn) Then
                        mCnn.BeginTrans
                            If Not (Rec.EOF Or Rec.BOF) Then
                                With mVoucher
                                    If val(VSGridZonalDetails.TextMatrix(mLoop, 11)) > 0 Then
                                        .intVoucherID_1 = val(VSGridZonalDetails.TextMatrix(mLoop, 11))
                                    Else
                                        .intVoucherID_1 = -1
                                    End If
                                    .intLocalBodyID_2 = Rec!intLocalBodyID
                                    .intTransactionID_3 = Rec!intTransactionID
                                    .intTransactionTypeID_4 = Rec!intTransactionTypeID
                                    .tnyVoucherTypeID_5 = Rec!tnyVoucherTypeID
                                    .intVoucherNo_6 = Rec!intVoucherNo
                                    .intBookNo_7 = Rec!intBookNo
                                    .dtDate_8 = Rec!dtDate
                                    .fltAmount_9 = Rec!fltAmount
                                    .intInstrumentTypeID_10 = Rec!intInstrumentTypeID
                                    .vchInstrumentNo_11 = Rec!vchInstrumentNo
                                    .dtInstrumentDate_12 = Rec!dtInstrumentDate
                                    .vchDescription_13 = Rec!vchDescription
                                    .numZoneID_14 = Rec!numZoneID
                                    .numWardID_15 = Rec!numWardId
                                    .intDoorNoP1_16 = Rec!intDoorNoP1
                                    .vchDoorNoP2_17 = Rec!vchDoorNoP2
                                    .vchDoorNoP3_18 = Rec!vchDoorNoP3
                                    .intUserID_19 = Rec!intUserID
                                    .intCounterID_20 = Rec!intCounterID
                                    .numSubLedgerID_21 = Rec!numSubLedgerID
                                    .intKeyID1_22 = Rec!intKeyID1
                                    .intKeyID2_23 = Rec!intKeyID2
                                    .intExternalApplicationID_24 = Rec!intExternalApplicationID
                                    .intExternalModuleID_25 = Rec!intExternalModuleID
                                    .intFinancialYearID_26 = Rec!intFinancialYearID
                                    .tnyShiftID_27 = Rec!tnyShiftID
                                    .tnyPrintFlag_28 = Rec!tnyPrintFlag
                                    .tnyCancelFlag_29 = Rec!tnyCancelFlag
                                    .dtRealisationDate = Rec!dtRealisationDate
                                    .vchRemarks = Rec!vchRemarks
                                    .vchBank_33 = Rec!vchBank
                                    .vchBankPlace_34 = Rec!vchBankPlace
                                    .intFundID_35 = Rec!intFundID
                                    .numSeatID = Rec!numSeatID
                                    .intSessionID = Rec!intSessionID
                                    .vchRefNo = Rec!vchRefNo
                                    .fltRoundOff = Rec!fltRoundOff
                                    .tnyReconciled = Rec!tnyReconciled
                                    .numTockenID = Rec!numTockenID
                                    .fltAdvAmtAdj = Rec!fltAdvAmtAdj
                                    .numInwardNo = Rec!numInwardNo
                                    .tnyStatus_32 = Rec!tnyStatus
                                    .numLocationID = Rec!numLocationID
                                    .tnyVoucherGroupID = Rec!tnyVoucherGroupID
                                    .numLinkKeyID = Rec!numLinkKeyID
                                    .dtTimeStamp = Rec!dtTimeStamp
                                    .dtChequeRealiseDate = Rec!dtChequeRealiseDate
                                    .vchVersionKey = Rec!vchVersionKey
                                    .tnyReversed = Rec!tnyReversed
                                    .dtValueDate = Rec!dtValueDate
                                 
                                    aryIn1 = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
                                    .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
                                    .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
                                    .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
                                    .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
                                    .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
                                    .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
                                    .tnyPrintFlag_28, .tnyCancelFlag_29, .dtRealisationDate, .vchRemarks, _
                                    .tnyStatus_32, .vchBank_33, .vchBankPlace_34, _
                                    .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
                                    .fltRoundOff, .tnyReconciled, .numTockenID, .fltAdvAmtAdj, .numInwardNo, _
                                    .numLocationID, .tnyVoucherGroupID, .numLinkKeyID, .dtTimeStamp, _
                                    .dtChequeRealiseDate, .vchVersionKey, .tnyReversed, .dtValueDate)
                                End With
                                Rec.Close
                                aryOut = ""
                                objdb.ExecuteSP "spZonalInsertVouchers", aryIn1, aryOut, , mCnn
                                mVoucherID = aryOut(0, 0)
                                mSql = "UPDATE faSyncLog SET tnyVouchers=2,intVoucherIDNew=" & mVoucherID & " WHERE intVoucherID=" & VSGridZonalDetails.TextMatrix(mLoop, 9) & " AND intLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                            End If
                            
                                
                                '-----------------VoucherChild---------------------------
                                
                                aryIn1 = ""
                                aryOut = ""
                                mSql = ""
                                mSql = "SELECT * FROM faVoucherChild WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND numLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                                If IsNumeric(val(VSGridZonalDetails.TextMatrix(mLoop, 11))) Then
                                    mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & val(VSGridZonalDetails.TextMatrix(mLoop, 11))
                                End If
            
                                mLoopHO = 0
                                While Not Rec.EOF Or Rec.BOF
                                    With mVouChildTbl
                                        .intVoucherID_1 = mVoucherID
                                        .intLocalBodyID_2 = Rec!intLocalBodyID
                                        .intSlNo_3 = Rec!intSlNo
                                        .intAccountHeadID_4 = Rec!intAccountHeadID
                                        .tnyDebitOrCredit_5 = Rec!tnyDebitOrCredit
                                        .intYearID_6 = Rec!intYearID
                                        .tnyPeriodID_7 = Rec!tnyPeriodID
                                        .tnyArrearFlag_8 = Rec!tnyArrearFlag
                                        .numDemandID_9 = Rec!numDemandID
                                        .fltAmount_10 = Rec!fltAmount
                                        
                                        aryIn1 = Array(.intVoucherID_1, _
                                                        .intLocalBodyID_2, _
                                                        .intSlNo_3, _
                                                        .intAccountHeadID_4, _
                                                        .tnyDebitOrCredit_5, _
                                                        .intYearID_6, _
                                                        .tnyPeriodID_7, _
                                                        .tnyArrearFlag_8, _
                                                        .numDemandID_9, _
                                                        .fltAmount_10)
            
                                        objdb.ExecuteSP "spSaveVoucherChild", aryIn1, , , mCnn
                                    End With
                                    aryIn1 = Null
                                    Rec.MoveNext
                                Wend
                                Rec.Close
                                mSql = ""
                                mSql = "UPDATE faSyncLog SET tnyVoucherChild=2 WHERE intVoucherID=" & VSGridZonalDetails.TextMatrix(mLoop, 9) & " AND intLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                            'End If
                            ''added on 11 Jan 2017 To Avoid Contra save in Vr Address and Vr Sub Tables
                            'If VSGridZonalDetails.TextMatrix(mRowCount, 13) <> 4001 Then
                                '---------------VoucherAddress--------------------------
                                aryIn1 = Null
                                aryOut = Null
                                mSql = ""
                                mSql = "SELECT * FROM faVoucherAddress WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND numLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                                If IsNumeric(val(VSGridZonalDetails.TextMatrix(mLoop, 11))) Then
                                    mCnn.Execute "Delete From faVoucherAddress Where intVoucherID = " & val(VSGridZonalDetails.TextMatrix(mLoop, 11))
                                End If
                                If Not Rec.EOF Or Rec.BOF Then
                                    With mVouAddress
                                        .intVoucherID = mVoucherID
                                        .intLocalBodyID = Rec!intLocalBodyID
                                        .vchName = Rec!vchName
                                        .vchHouseName = Rec!vchHouseName
                                        .vchStreetName = Rec!vchStreetName
                                        .vchMainPlace = Rec!vchMainPlace
                                        .vchPostOffice = Rec!vchPostOffice
                                        .vchDistrict = Rec!vchDistrict
                                        .vchPinNumber = Rec!vchPinNumber
                                        .vchInit1 = Rec!vchInit1
                                        .vchInit2 = Rec!vchInit2
                                        .vchInit3 = Rec!vchInit3
                                        .vchInit4 = Rec!vchInit4
                                        .vchLocalPlace = Rec!vchLocalPlace
                                        .vchPhone = Rec!vchPhone
                                        .intWardNo = Rec!intWardNo
                                        .intDoorNo = Rec!intDoorNo
                                        .vchDoorNo2 = Rec!vchDoorNo2
                                        aryIn1 = Array(.intVoucherID, _
                                                        .intLocalBodyID, _
                                                        .vchName, _
                                                        .vchInit1, _
                                                        .vchInit2, _
                                                        .vchInit3, _
                                                        .vchInit4, _
                                                        .vchHouseName, _
                                                        .vchStreetName, _
                                                        .vchLocalPlace, _
                                                        .vchMainPlace, _
                                                        .vchPostOffice, _
                                                        .vchDistrict, _
                                                        .vchPinNumber, _
                                                        .vchPhone, _
                                                        .intWardNo, _
                                                        .intDoorNo, _
                                                        .vchDoorNo2)
                                        objdb.ExecuteSP "spSaveVoucherAddress", aryIn1, , , mCnn
                                    End With
                                    Rec.Close
                                    mSql = "UPDATE faSyncLog SET tnyVoucherAddress=2 WHERE intVoucherID=" & VSGridZonalDetails.TextMatrix(mLoop, 9) & " AND intLocationID=" & ZonalDetailsID & ""
                                    Rec.Open mSql, mCnnHO
                                End If
                                
                                '----------------VoucherSub------------------------------
                                aryIn1 = Null
                                aryOut = Null
                                mSql = ""
                                mSql = "SELECT * FROM faVoucherSub WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND numLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                                If Not Rec.EOF Or Rec.BOF Then
                                    With mVoucherSub
                                        .intVoucherID = mVoucherID
                                        .intLocalBodyID = Rec!intLocalBodyID
                                        .decProjectID = Rec!decProjectID
                                        .intSourceOfFundID = Rec!intSourceOfFundID
                                        .intCategoryID = Rec!intCategoryID
                                        .intSectorID = Rec!intSectorID
                                        .intAllotmentID = Rec!intAllotmentID
                                        .intAgreementID = Rec!intAgreementID
                                        .intCashBookID = Rec!intCashBookID
                                        .intImplementingOfficerID = Rec!intImplementingOfficerID
                                        .intCreditorTypeID = Rec!intCreditorTypeID
                                        .intCreditorsID = Rec!intCreditorsID
                                        .intTypeID = Rec!intTypeID
                                        aryIn1 = Array(.intVoucherID, _
                                                        .intLocalBodyID, _
                                                        .decProjectID, _
                                                        .intSourceOfFundID, _
                                                        .intCategoryID, _
                                                        .intSectorID, _
                                                        .intAllotmentID, _
                                                        .intAgreementID, _
                                                        .intCashBookID, _
                                                        .intImplementingOfficerID, _
                                                        .intCreditorTypeID, _
                                                        .intCreditorsID, _
                                                        .intTypeID)
                                        objdb.ExecuteSP "spSaveVoucherSub", aryIn1, , , mCnn
                                    End With
                                    Rec.Close
                                    mSql = "UPDATE faSyncLog SET tnyVoucherSub=2 WHERE intVoucherID=" & VSGridZonalDetails.TextMatrix(mLoop, 9) & " AND intLocationID=" & ZonalDetailsID & ""
                                    Rec.Open mSql, mCnnHO
                                End If
                            'End If
                            '----------------Transactions---------------------------
                            aryIn1 = Null
                            aryOut = Null
                            
                            mSql = ""
                            mSql = "SELECT intTransactionID FROM faTransactions WHERE intVoucherID=" & mVoucherID
                            Rec.Open mSql, mCnn
                            mTransactionID = -1
                            While Not Rec.EOF
                                mTransactionID = Rec!intTransactionID
                                Rec.MoveNext
                            Wend
                            Rec.Close
                            mSql = ""
                            mSql = "SELECT * FROM faTransactions WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND numLocationID=" & ZonalDetailsID & ""
                            Rec.Open mSql, mCnnHO
                            If Not Rec.EOF Or Rec.BOF Then
                                With mTranTable
                                    .intTransactionID = mTransactionID
                                    .intLocalBodyID = Rec!intLocalBodyID
                                    .intFinancialYearID = Rec!intFinancialYearID
                                    .dtTransactionDate = Rec!dtTransactionDate
                                    .intExternalApplicationID = Rec!intExternalApplicationID
                                    .intExternalApplicationModuleID = Rec!intExternalApplicationModuleID
                                    .intFunctionID = Rec!intFunctionID
                                    .intFunctionaryID = Rec!intFunctionaryID
                                    .intFieldID = Rec!intFieldID
                                    .intFundID = Rec!intFundID
                                    .intBudgetCentreID = Rec!intBudgetCentreID
                                    .vchNarration = Rec!vchNarration
                                    .intTransactionTypeID = Rec!intTransactionTypeID
                                    .intProcessID = Rec!intProcessID
                                    .vchGroup = Rec!vchGroup
                                    .intGroupID = Rec!intGroupID
                                    .intKeyID = Rec!intKeyID
                                    .numSubLedgerID = Rec!numSubLedgerID
                                    .numUserID = Rec!numUserID
                                    .intVoucherID = mVoucherID
                                    .tnyVoucherGroupID = Rec!tnyVoucherGroupID
                                    .intVoucherNo = Rec!intVoucherNo
                                    .tnyStatus = Rec!tnyStatus
                                    .tnyReversed = Rec!tnyReversed
                                    .dtValueDate = Rec!dtValueDate
                                
                                    
                                    aryIn1 = Array(.intTransactionID, _
                                                    .intLocalBodyID, _
                                                    .intFinancialYearID, _
                                                    .dtTransactionDate, _
                                                    .intExternalApplicationID, _
                                                    .intExternalApplicationModuleID, _
                                                    .intFunctionID, _
                                                    .intFunctionaryID, _
                                                    .intFieldID, _
                                                    .intFundID, _
                                                    .intBudgetCentreID, _
                                                    .vchNarration, _
                                                    .intTransactionTypeID, _
                                                    .intProcessID, _
                                                    .vchGroup, _
                                                    .intGroupID, _
                                                    .intKeyID, _
                                                    .numSubLedgerID, _
                                                    .numUserID, _
                                                    .intVoucherID, _
                                                    .tnyVoucherGroupID, _
                                                    .intVoucherNo, _
                                                    .tnyStatus, _
                                                    .tnyReversed, _
                                                    .dtValueDate)
                                    aryOut = Null
                                    objdb.ExecuteSP "spSaveTransactions", aryIn1, aryOut, , mCnn
                                End With
                                Rec.Close
                                mTransactionID = aryOut(0, 0)
                                mSql = ""
                                mSql = "UPDATE faSyncLog SET tnyTransactions=2 WHERE intVoucherID=" & VSGridZonalDetails.TextMatrix(mLoop, 9) & " AND intLocationID=" & ZonalDetailsID & ""
                                Rec.Open mSql, mCnnHO
                            End If
                            
                            '--------------TransactionChild---------------------------
                            aryIn1 = Null
                            aryOut = Null
                            mSql = ""
                            mSql = "SELECT faTransactionChild.fltAmount ChdAmt,* FROM faTransactionChild INNER JOIN faTransactions ON "
                            mSql = mSql + "faTransactionChild.intTransactionID=faTransactions.intTransactionID "
                            mSql = mSql + " AND faTransactionChild.numLocationID=faTransactions.numLocationID "
                            mSql = mSql + " INNER JOIN faVouchers ON faTransactions.intVoucherID=faVouchers.intVoucherID "
                            mSql = mSql + " AND faTransactions.numLocationID=faVouchers.numLocationID "
                            mSql = mSql + "WHERE faTransactions.intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND faTransactions.numLocationID=" & ZonalDetailsID & ""
                            Rec.Open mSql, mCnnHO
                            If IsNumeric(val(VSGridZonalDetails.TextMatrix(mLoop, 11))) Then
                                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mTransactionID
                            End If
                            mLoopHO = 0
                            While Not Rec.EOF Or Rec.BOF
                                mSql = ""
                                With mTranChildTbl
                                    .intTransactionID = mTransactionID
                                    .intSerialNo = Rec!intSerialNo
                                    .intAccountHeadID = Rec!intAccountHeadID
                                    .fltAmount = Rec!ChdAmt
                                    .tinDebitOrCreditFlag = Rec!tinDebitOrCreditFlag
                                    .intByAccountHeadID = Rec!intByAccountHeadID
                                    .vchNarration = Rec!vchNarration
                                    .intFundID = Rec!intFundID
                                    .fltOpeningBalance = Rec!fltOpeningBalance
                                    .numTockenID = Rec!numTockenID
                                    .dtReconcileDate = Rec!dtReconcileDate
                                    
                                    aryIn1 = Array(.intTransactionID, _
                                                    .intSerialNo, _
                                                    .intAccountHeadID, _
                                                    .fltAmount, _
                                                    .tinDebitOrCreditFlag, _
                                                    .intByAccountHeadID, _
                                                    .vchNarration, _
                                                    .intFundID)
                                                    ', _
                                                    '.fltOpeningBalance, _
                                                    '.numTockenID, _
                                                    '.dtReconcileDate)
                                    
                                    objdb.ExecuteSP "spSaveTransactionChild", aryIn1, , , mCnn
                                End With
                                aryIn1 = Null
                                Rec.MoveNext
                            Wend
                            Rec.Close
                            mSql = "UPDATE faSyncLog SET tnyTransactionChild=2 WHERE intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND intLocationID=" & ZonalDetailsID & ""
                            Rec.Open mSql, mCnnHO
                            mSql = ""
                            mSql = "UPDATE faSyncLog Set tnySyncStatus=2 WHERE tnyVouchers=2 AND tnyVoucherChild=2 AND tnyVoucherSub=2 AND tnyVoucherAddress=2 AND tnyTransactions=2"
                            mSql = mSql + " AND tnyTransactionChild=2 AND intVoucherID=" & val(VSGridZonalDetails.TextMatrix(mLoop, 9)) & " AND intLocationID=" & ZonalDetailsID & ""
                            Rec.Open mSql, mCnnHO
                            
                            mCnnHO.Close
                            
                            VSGridZonalDetails.Cell(flexcpChecked, mLoop, 8) = vbChecked
                            'VSGridZonalDetails.Cell(flexcpBackColor, mLoop, 0, mLoop, 11) = &HC0FFFF
                        mCnn.CommitTrans
                    Else
                        MsgBox "Connection failed", vbInformation, "Saankhya"
                    'mCnn.Close
                    End If
                End If
            End If
        Next mLoop
        If mLoop = VSGridZonalDetails.Rows Then
            MsgBox "Data Transfer Successfully Completed", vbInformation, "Saankhya"
            cmdTransafer.Enabled = False
        Else
            MsgBox "Data Transfer Not Completed", vbInformation, "Saankhya"
        End If
    End Sub

    Private Sub cmdVerify_Click()
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim aryIn As Variant
        Dim intFinYearID As Integer
        
        If Not IsDate(mdtZonDate) Then
            MsgBox "Date not set", vbInformation
            Exit Sub
        End If
        
        intFinYearID = GetYearID(CDate(mdtZonDate))
        
        If Not objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO) Then
            MsgBox "Connection to HO not found", vbInformation
            Exit Sub
        End If
        cmdVerify.Enabled = False
        mSql = ""
        mSql = "SELECT faSyncLog.intVoucherID FROM faSyncLog INNER JOIN faVouchers "
        mSql = mSql + "ON  faVouchers.intVoucherID=faSyncLog.intVoucherID AND faVouchers.numLocationID=faSyncLog.intLocationID"
        mSql = mSql + " WHERE tnySyncStatus<>2 AND dtDate < '" & DdMmmYy(txtDate.Text) & "'"
        Rec.Open mSql, mCnn
        If Rec.BOF Or Rec.EOF Then
            Me.MousePointer = vbHourglass
            aryIn = Array(mnumZonDetailsID, mdtZonDate, intFinYearID)
            objdb.ExecuteSP "spZonalVoucherVerify", aryIn, , , mCnn, adCmdStoredProc
            
            chkWrongVouchers.Value = 1
            Call FillGrid
            
            For mLoop = 1 To VSGridZonalDetails.Rows - 1
                If VSGridZonalDetails.RowHidden(mLoop) = False Then
                    Exit For
                End If
            Next mLoop
            If mLoop = VSGridZonalDetails.Rows Then
                chkWrongVouchers.Value = 0
                Call FillGrid
            End If
            Me.MousePointer = vbDefault
        Else
            MsgBox "Verification is not Possible,Please Transfer Previous  Vouchers", vbInformation
        End If
        '        mFlag = 0
        '        For mLoop = 1 To VSGridZonalDetails.Rows - 1
        '            aryIn = Array(val(VSGridZonalDetails.TextMatrix(mLoop, 9)), mnumZonDetailsID)
        '            Set Rec = objDB.ExecuteSP("spZonalVoucherVerify", aryIn, , , mCnn, adCmdStoredProc)
        '            If Not (Rec.EOF And Rec.BOF) Then
        '                VSGridZonalDetails.Cell(flexcpChecked, mLoop, 7) = vbChecked
        '                'VSGridZonalDetails.Cell(flexcpBackColor, mLoop, 0, mLoop, 11) = &HC0FFC0
        '            Else
        '                mFlag = 1
        '
        '                VSGridZonalDetails.TextMatrix(mLoop, 10) = "Incorrect Voucher"
        '                VSGridZonalDetails.Cell(flexcpBackColor, mLoop, 0, mLoop, 11) = &H8080FF
        '            End If
        '        Next mLoop
        '        If mFlag = 0 Then
        '            MsgBox "Verification completed"
        '        Else
        '            MsgBox "Verification not Completed"
        '        End If
        '        Rec.Close
        '        mCnn.Close
    End Sub
    
Private Sub Command2_Click()

End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = frmMenu.Top + 1080
        Call FormIntialize
        Call FillGrid
        Call GetContraAmt
       ' Call VerifyContraAmt
    End Sub
    
    Private Sub FormIntialize()
        If IsDate(mdtZonDate) Then
            txtDate.Text = DdMmmYy(CDate(mdtZonDate))
        End If
        lblZonal.Caption = mvarZonName
        cmdTransafer.Enabled = False
        cmdVerify.Enabled = False
'                Dim objDB As New clsDB
'        Dim mCnn As New ADODB.Connection
'        Dim mSQL As String
'        Dim Rec As New Recordset
'        Dim mRowCount As Variant
'        Dim aryIn(1) As Variant
        
'        aryIn(0) = DdMmmYy(txtDate.Text)
'        aryIn(1) = ZonalDetailsID
'        objDB.CreateNewConnection mCnn, SaankhyaHO
'        Set Rec = objDB.ExecuteSP("spGetVouchersData", aryIn, , , mCnn)
'        mRowCount = 1
'        VSGridZonalDetails.Clear 1, 1
'        VSGridZonalDetails.Rows = 1
'        While Not Rec.EOF
'            VSGridZonalDetails.Rows = VSGridZonalDetails.Rows + 1
'            VSGridZonalDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'            VSGridZonalDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'            VSGridZonalDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
'            VSGridZonalDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'            VSGridZonalDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!cashAmt), "", Rec!cashAmt)
'            VSGridZonalDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!BankAmt), "", Rec!BankAmt)
'            VSGridZonalDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo) + " " + IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
'            VSGridZonalDetails.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
'            VSGridZonalDetails.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!intVoucherIDNew), "", Rec!intVoucherIDNew)
'            If Not IsNull(Rec!tnySyncStatus) Then
'                If (Rec!tnySyncStatus = 2) Then
'                    VSGridZonalDetails.TextMatrix(mRowCount, 7) = vbChecked
'                    VSGridZonalDetails.TextMatrix(mRowCount, 8) = vbChecked
'                End If
'            End If
'            Rec.MoveNext
'            mRowCount = mRowCount + 1
'        Wend
'        Rec.Close
'        Call Calculation
'        mCnn.Close
    End Sub
    
    Private Sub Calculation()
        Dim mLoop As Integer
        Dim mCashTotal As Double, mBankTotal As Double
        mCashTotal = 0
        mBankTotal = 0
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim mRowCount As Variant
'        Dim aryIn(2) As Variant
'
'        aryIn(0) = mnumZonDetailsID
'        aryIn(1) = gbFinancialYearID
'        aryIn(2) = mdtZonDate
'        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        Set Rec = objdb.ExecuteSP("spgetzonalcontra", aryIn, , , mCnn)
        
        For mLoop = 1 To VSGridZonalDetails.Rows - 1  '' Added By anisha On 18.Sep.2014
            If val(VSGridZonalDetails.TextMatrix(mLoop, 12)) = 0 Then
                mCashTotal = mCashTotal + val(VSGridZonalDetails.TextMatrix(mLoop, 4))
                mBankTotal = mBankTotal + val(VSGridZonalDetails.TextMatrix(mLoop, 5))
            End If
        Next

        txtTotalCash.Text = Format(mCashTotal, "0.00")
        txtTotalBankAmt.Text = Format(mBankTotal, "0.00")
        txtContraAmount = Format(mCashTotal, "0.00")
    End Sub
    
    Public Property Let ZonalDetailsID(mData As Variant)
        mnumZonDetailsID = mData
    End Property
    
    Public Property Get ZonalDetailsID() As Variant
        ZonalDetailsID = mnumZonDetailsID
    End Property
    
    Public Property Let ZonalName(mData As Variant)
        mvarZonName = mData
    End Property
    
    Public Property Get ZonalName() As Variant
        ZonalName = mvarZonName
    End Property
    
    Public Property Let ZonalTrnDate(mData As Variant)
        mdtZonDate = mData
    End Property
    
    Public Property Get ZonalTrnDate() As Variant
        ZonalTrnDate = mdtZonDate
    End Property

Private Sub txtAccountCode_GotFocus()

    If gbSearchStr <> "" Then
            Dim mStr As String
            txtAccountCode.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            
        End If
        txtAccountCode.SelStart = 0
        txtAccountCode.SelLength = Len(txtAccountCode)
End Sub


