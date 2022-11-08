VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtractYearWiseList 
   Caption         =   "Extract YearWise List"
   ClientHeight    =   8520
   ClientLeft      =   75
   ClientTop       =   255
   ClientWidth     =   14790
   Icon            =   "frmExtractYearWiseList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14790
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   15525
      ScaleHeight     =   7575
      ScaleWidth      =   3435
      TabIndex        =   8
      Top             =   945
      Width           =   3435
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PUBLISH the Data"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   450
         TabIndex        =   18
         Top             =   5235
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   5220
         Width           =   195
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enter the Council resolution no. And Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   870
         Left            =   450
         TabIndex        =   16
         Top             =   4020
         Width           =   2760
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   4005
         Width           =   195
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Verify the Assets And Liability balances"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   780
         Left            =   450
         TabIndex        =   14
         Top             =   2805
         Width           =   2985
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   2790
         Width           =   195
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Verify the Cash/Bank/tresury Closing Balances"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   930
         Left            =   450
         TabIndex        =   12
         Top             =   1485
         Width           =   2805
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1485
         Width           =   195
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Extracts the Yearly/Monthly Balances"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   540
         TabIndex        =   10
         Top             =   285
         Width           =   2760
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   270
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   18990
      TabIndex        =   7
      Top             =   9990
      Width           =   19050
   End
   Begin VB.Frame frmProgressBar 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   45
      TabIndex        =   4
      Top             =   8550
      Visible         =   0   'False
      Width           =   14595
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   180
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblExtracting 
         Alignment       =   2  'Center
         Caption         =   "EXTRACTING!!!!!!!!!!!!!!!"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         TabIndex        =   6
         Top             =   585
         Width           =   4470
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   19950
      TabIndex        =   1
      Top             =   45
      Width           =   19950
      Begin VB.Label lblHeadings 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   45
         TabIndex        =   3
         Top             =   135
         Width           =   8160
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   7620
      Left            =   1800
      TabIndex        =   0
      Top             =   900
      Width           =   13560
      _cx             =   23918
      _cy             =   13441
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
      Rows            =   5
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExtractYearWiseList.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
   Begin VB.Image Image1 
      Height          =   7560
      Left            =   45
      Picture         =   "frmExtractYearWiseList.frx":1EC1
      Stretch         =   -1  'True
      Top             =   900
      Width           =   1755
   End
   Begin VB.Label Label1 
      Height          =   8565
      Left            =   15390
      TabIndex        =   2
      Top             =   900
      Width           =   3660
   End
End
Attribute VB_Name = "frmExtractYearWiseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mStatus   As Integer
Dim mModuleID As Integer

Private Sub Form_Activate()
    frmProgressBar.Visible = False
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    
    vsGrid.MergeRow(0) = True
    vsGrid.MergeCol(0) = True
    vsGrid.MergeCol(1) = True
    vsGrid.MergeCol(2) = True
    vsGrid.MergeCol(4) = True
    vsGrid.MergeCol(5) = True
    
    lblHeadings.Caption = " CLOSING YEARWISE ACCOUNTS"
    Call FillGrid
    
    frmProgressBar.Visible = False
End Sub


Private Sub vsGrid_DblClick()
    Dim mLoop As Integer
    If vsGrid.Row > 0 Then
        
        
        For mLoop = 2 To vsGrid.Row - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 5) = 2 And vsGrid.Row <> 1 Then  '(vsGrid.Cell(flexcpChecked, mLoop, 1) = 2 Or vsGrid.Cell(flexcpChecked, mLoop, 2) = 2) And vsGrid.Row <> 1
                MsgBox "Verify the Previous Year's Data"
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        Next mLoop
        
        If vsGrid.Col = 1 Then
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 1) = 2 And val(vsGrid.TextMatrix(vsGrid.Row, 6)) < 2014 Then
                Me.MousePointer = vbHourglass
                Call ExtarctDailyExtract(val(vsGrid.TextMatrix(vsGrid.Row, 6)))
                Me.MousePointer = vbDefault
                frmExtractedCashBook.LoadMode = 1
                frmExtractedCashBook.YearID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                frmExtractedCashBook.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                frmExtractedCashBook.Show vbModal
                
                Exit Sub
            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 6)) < 2014 Then
               frmExtractedCashBook.LoadMode = 1
               frmExtractedCashBook.YearID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
               frmExtractedCashBook.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
               frmExtractedCashBook.cmdVerify.Enabled = False
               frmExtractedCashBook.Show vbModal
            End If
        End If
        If vsGrid.Col = 2 Then
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 1) = 1 Then
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 2) = 2 And val(vsGrid.TextMatrix(vsGrid.Row, 6)) < 2014 Then
                    frmExtractedBalanceSheet.LoadMode = 1
                    frmExtractedBalanceSheet.YearID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                    frmExtractedBalanceSheet.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                    frmExtractedBalanceSheet.Show vbModal
                    Me.MousePointer = vbDefault
                    Exit Sub
                ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 6)) < 2014 Then
                    frmExtractedBalanceSheet.LoadMode = 1
                    frmExtractedBalanceSheet.YearID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                    frmExtractedBalanceSheet.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                    frmExtractedBalanceSheet.cmdVerify.Enabled = False
                    If val(vsGrid.TextMatrix(vsGrid.Row, 7)) = 2 Then
                        frmExtractedBalanceSheet.cmdPublish.Enabled = False
                        frmExtractedBalanceSheet.txtCouncilNo.Text = val(vsGrid.TextMatrix(vsGrid.Row, 3))
                        frmExtractedBalanceSheet.txtCouncilDate.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
                    Else
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            frmExtractedBalanceSheet.cmdPublish.Enabled = True
                        Else
                            frmExtractedBalanceSheet.cmdPublish.Enabled = False
                        End If
                    End If
                    frmExtractedBalanceSheet.Show vbModal
                    
                End If
            Else
                MsgBox "CASH AND BANK BALANCES NOT VERIFIED YET!!!!!!!!", vbInformation
                Exit Sub
            End If
        End If
        
        If val(vsGrid.TextMatrix(vsGrid.Row, 6)) >= 2014 Then
            'frmExtactedMonthWiseList.YearID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
            'frmExtactedMonthWiseList.Show vbModal
            MsgBox "Year 2014 Extraction is not Enabled!", vbInformation
            
            Exit Sub
        End If
        'Unload Me
        
    End If
    Me.MousePointer = vbDefault
End Sub
Public Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mDate As Date
    Dim mYearID As Variant
    Dim mCount, i As Integer

    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " SELECT   dtDate, intFinancialYearID mYear FROM faVouchers  WHERE intTransactionTypeID=3000"
    
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        'mDate = DdMmmYy(Rec!dtDate)
        mYearID = Rec!mYear
    End If
    Rec.Close
    
'''    If CDate(mDate) >= CDate("01/Apr/2009") And CDate(mDate) <= CDate("31/Mar/2010") Then
'''        mYearID = 2009
'''    ElseIf CDate(mDate) >= CDate("01/Apr/2010") And CDate(mDate) <= CDate("31/Mar/2011") Then
'''        mYearID = 2010
'''    ElseIf CDate(mDate) >= CDate("01/Apr/2011") And CDate(mDate) <= CDate("31/Mar/2012") Then
'''        mYearID = 2011
'''    ElseIf CDate(mDate) >= CDate("01/Apr/2012") And CDate(mDate) <= CDate("31/Mar/2013") Then
'''        mYearID = 2012
'''    End If
    
    mCount = 2
    vsGrid.Clear 1, 1
    vsGrid.Rows = 2
    mYearID = mYearID + 1
        
    If mYearID = gbFinancialYearID Then
        vsGrid.TextMatrix(2, 0) = CStr(mYearID) + "-" + mID$(CStr(mYearID + 1), 3, 2)
        vsGrid.TextMatrix(2, 6) = mYearID
    Else
        For i = mYearID To gbFinancialYearID
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.TextMatrix(mCount, 0) = CStr(i) + "-" + mID$(CStr(i + 1), 3, 2)
            vsGrid.TextMatrix(mCount, 6) = i
            mCount = mCount + 1
            
        Next i
    End If
    Call VerifyStatus
End Sub
Private Sub VerifyStatus()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mRowCnt = 1
    mSQL = " SELECT * FROM faPostingIndex "
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then

        'vsGrid.Row = vsGrid.Row + 1
        While Not (Rec.EOF Or Rec.BOF)
            If val(vsGrid.TextMatrix(mRowCnt, 6)) = Rec!intFinYearID And val(vsGrid.TextMatrix(mRowCnt, 6)) < 2014 Then
                If Rec!tnyVerifyCash = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 1) = vbChecked
                End If
                If Rec!tnyVerifyBS = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 2) = vbChecked
                End If
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                If Rec!dtRefDate <> "" Then
                    vsGrid.TextMatrix(mRowCnt, 4) = DdMmmYy(IIf(IsNull(Rec!dtRefDate), "", Rec!dtRefDate))
                Else
                    vsGrid.TextMatrix(mRowCnt, 4) = ""
                End If
                If Rec!tnyStage = 2 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 5) = vbChecked
                End If
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!tnyStage), "", Rec!tnyStage)
                vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intMonthID), "", Rec!intMonthID)
            
            ElseIf val(vsGrid.TextMatrix(mRowCnt, 6)) >= 2014 Then
                If val(vsGrid.TextMatrix(mRowCnt, 6)) = Rec!intFinYearID And Rec!intMonthID = 3 Then
                    If Rec!tnyVerifyCash = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 1) = vbChecked
                End If
                If Rec!tnyVerifyBS = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 2) = vbChecked
                End If
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                If Rec!dtRefDate <> "" Then
                    vsGrid.TextMatrix(mRowCnt, 4) = DdMmmYy(IIf(IsNull(Rec!dtRefDate), "", Rec!dtRefDate))
                Else
                    vsGrid.TextMatrix(mRowCnt, 4) = ""
                End If
                If Rec!tnyStage = 2 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 5) = vbChecked
                End If
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!tnyStage), "", Rec!tnyStage)
                vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intMonthID), "", Rec!intMonthID)
                
                End If
               
            End If
            
        mRowCnt = mRowCnt + 1
        Rec.MoveNext
        If mRowCnt = vsGrid.Rows Then Exit Sub
    Wend
    Rec.Close
    End If
End Sub
Private Sub ExtarctDailyExtract(mYearID As Integer)
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec  As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    Dim mExtractFlag As Boolean
    Dim mArrIn As Variant
    Dim mSqlChild As String
    Dim mCount As Long
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    ProgressBar1.value = 0
    
    mRowCnt = 1
    mSQL = " SELECT * FROM faDailyExtracts WHERE   intTypeID<>0 AND intFinancialYearID=" & mYearID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mExtractFlag = True
    End If
    If mExtractFlag = False Then
        mSqlChild = "SELECT dtDate,intMonthID,intFinancialYearID,intTypeID,intFunctionID,intSourceOfFundID,intAccountHeadID,SUM(fltAmount) fltAmount,intLocalBodyID FROM ("
        mSqlChild = mSqlChild + "  SELECT CASE"
        mSqlChild = mSqlChild + "      WHEN faTransactions.tnyReversed = 1 AND intGroupID = 20  THEN 11"
        mSqlChild = mSqlChild + "      WHEN faTransactions.tnyReversed = 1 AND intGroupID = 10  THEN 21"
        mSqlChild = mSqlChild + "      WHEN faTransactions.tnyVoucherGroupID = 2 AND intGroupID = 40 AND LEFT(ISNULL(numLinkKeyID,0),1) <> 2 THEN 12"
        mSqlChild = mSqlChild + "      WHEN faTransactions.tnyVoucherGroupID = 2 AND intGroupID = 40 AND LEFT(ISNULL(numLinkKeyID,0),1) = 2 THEN 22"
        mSqlChild = mSqlChild + "      Else: intGroupID"
        mSqlChild = mSqlChild + "         END intTypeID, intFunctionID, intSourceOfFundID, intAccountHeadID, (tinDebitOrCreditFlag * -2 + 1) * -1 * faTransactionChild.fltAmount fltAmount, faTransactions.intFinancialYearID, Month(dtDate) intMonthID, dtDate, faTransactions.intLocalBodyID"
        mSqlChild = mSqlChild + "  From faTransactionChild"
        mSqlChild = mSqlChild + "  INNER JOIN faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID"
        mSqlChild = mSqlChild + "  INNER JOIN faVouchers ON faVouchers.intVoucherID = faTransactions.intVoucherID"
        mSqlChild = mSqlChild + "  LEFT JOIN faVoucherSub ON faVoucherSub.intVoucherID = faVouchers.intVoucherID"
        mSqlChild = mSqlChild + "  WHERE ISNULL(faVouchers.tnyCancelFlag,0) <> 1 AND faVouchers.intFinancialYearID = " & mYearID
        mSqlChild = mSqlChild + "  ) A"
        mSqlChild = mSqlChild + "  GROUP By dtDate, intMonthID, intFinancialYearID, intTypeID, intAccountHeadID,intFunctionID,intSourceOfFundID, intLocalBodyID"
        RecChild.CursorLocation = adUseClient
        RecChild.Open mSqlChild, mCnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RecChild.EOF And RecChild.BOF) Then
            mCount = RecChild.RecordCount
        End If
        RecChild.Close
        
        mArrIn = Array(mYearID)
        objDB.ExecuteSP "spDailyExtractByYear", mArrIn, , , mCnn, adCmdStoredProc
        
        If val(vsGrid.TextMatrix(2, 6)) = mYearID Then
            objDB.ExecuteSP "spDailyExtract_Opening", mArrIn, , , mCnn, adCmdStoredProc
        End If
        
        frmProgressBar.Visible = True

        ProgressBar1.Max = mCount + 1
        While ProgressBar1.value < ProgressBar1.Max
                ProgressBar1.value = ProgressBar1.value + 1
        Wend
    Else
        frmProgressBar.Visible = False
    End If
    Rec.Close
    mCnn.Close
End Sub

