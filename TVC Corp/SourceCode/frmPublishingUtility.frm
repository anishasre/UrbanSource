VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPublishingUtility 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PUBLISH STATUS"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "frmPublishingUtility.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   7080
      Width           =   1815
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5655
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   10575
      _cx             =   18653
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPublishingUtility.frx":1CCA
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
   Begin VB.CommandButton cmdSync 
      Caption         =   "<<SYNC>>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13455
      TabIndex        =   0
      Top             =   0
      Width           =   13455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "MONTH END CLOSING BALANCES"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
End
Attribute VB_Name = "frmPublishingUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mYearID As Integer
    Dim dtDate As Variant
    Dim mMonthID As Integer
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub Form_Load()
        Call SetSyncDate
        Call UpdateDailyIndex
    End Sub
    
    Private Sub SetSyncDate()
        Dim mCnn                As New ADODB.Connection
        Dim objDB               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim mSQl                As String
        Dim dtClosingDate       As Variant
        Dim mArrIn              As Variant
        
        '==================================================================================='
        'GET CURRENT SYNC YEAR AND MONTH FROM CONFIG TABLE
        '==================================================================================='
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSQl = "SELECT SynVerificationYearID,SynVerificationDate,Month(SynVerificationDate) mMonthID FROM faConfig"
            Rec.Open mSQl, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mYearID = IIf(IsNull(Rec!SynVerificationYearID), -1, Rec!SynVerificationYearID)
                dtDate = IIf(IsNull(Rec!SynVerificationDate), -1, Rec!SynVerificationDate)
                mMonthID = IIf(IsNull(Rec!mMonthID), -1, Rec!mMonthID)
            End If
            Rec.Close
            
            If mYearID = -1 Then
                mSQl = ""
                mSQl = " SELECT    dtDate,Month(dtDate) mMonthID, intFinancialYearID mYear FROM faVouchers  WHERE intTransactionTypeID=3000"
                Rec.Open mSQl, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mYearID = (Rec!mYear) + 1
                    dtDate = DdMmYy(Rec!dtDate)
                    mMonthID = Rec!mMonthID
                End If
                Rec.Close
                '==================================================================================='
                'UPDATE faCONFIG - SynVerificationYearID||SynVerificationDate
                '==================================================================================='
                 mSQl = ""
                 mSQl = "UPDATE faConfig SET SynVerificationYearID=" & mYearID & " ,SynVerificationDate=" & dtDate & "  "
                 objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
                
            End If
            
            
        '==================================================================================='
        ' CHECK VALIDATE CURRENT SYNC YEAR WITH POSTING INDEX TABLE
        ' IF YEAR NOT FOUND IN POSTING INDEX - INSERT YEAR IN POSTING INXEX
        '==================================================================================='
            mSQl = " SELECT * FROM faPostingIndex WHERE ISNULL(tnyVerifyCash,0)=1 AND ISNULL(tnyVerifyBS,0)=1 AND intFinYearID= " & mYearID
            Rec.Open mSQl, mCnn
            If (Rec.EOF And Rec.BOF) Then
                dtClosingDate = "31/Mar/" + CStr(mYearID)
                mArrIn = Array(-1, mYearID, 3, dtClosingDate, _
                                dtClosingDate, _
                                Null, _
                                Null, _
                                0, _
                                0, _
                                1, _
                                0 _
                                )
                objDB.ExecuteSP "spSavePostingIndexCashBank", mArrIn, , , mCnn, adCmdStoredProc
            End If
            Rec.Close
        '==================================================================================='
        txtYear.Text = mYearID
        txtMonth.Text = MonthName(mMonthID)
            
        mCnn.Close
    End Sub

    Private Sub FillGrid()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQl  As String
        Dim mRowCnt As Integer
    
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQl = " SELECT A.dtDate,SUM(CASHBALANCE) CASHBALANCE,SUM(BANKBALANCE) BANKBALANCE FROM"
        mSQl = mSQl + " ("
        mSQl = mSQl + " SELECT  faDailyExtracts.dtDate dtDate,"
        mSQl = mSQl + "         CASE WHEN faAccountHeads.intGroupID=1 Then"
        mSQl = mSQl + "         SUM(fltAmount) ELSE 0 END CASHBALANCE,"
        mSQl = mSQl + "         CASE WHEN faAccountHeads.intGroupID=2 Then"
        mSQl = mSQl + "         SUM(fltAmount) ELSE 0 END BANKBALANCE"
        mSQl = mSQl + " FROM faDailyExtracts"
        mSQl = mSQl + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faDailyExtracts.intAccountHeadID"
        mSQl = mSQl + " WHERE faDailyExtracts.intFinancialYearID = " & mYearID & ""
        mSQl = mSQl + " AND intMonthID = " & mMonthID & ""
        mSQl = mSQl + "       AND faAccountHeads.intGroupID in (1,2)"
        mSQl = mSQl + " GROUP BY faDailyExtracts.dtDate,faAccountHeads.intGroupID"
        mSQl = mSQl + " )A"
        mSQl = mSQl + " GROUP BY A.dtDate"
        mSQl = mSQl + " ORDER BY A.dtDate"
    
        Rec.CursorLocation = adUseClient
        Rec.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not (Rec.EOF Or Rec.BOF)
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.TextMatrix(mRowCnt, 0) = DdMmmYy(IIf(IsNull(Rec!dtDate), "", Rec!dtDate))
            vsGrid.TextMatrix(mRowCnt, 1) = Format(IIf(IsNull(Rec!CASHBALANCE), "", Rec!CASHBALANCE), "0.00")
            vsGrid.TextMatrix(mRowCnt, 2) = Format(IIf(IsNull(Rec!BankBalance), "", Rec!BankBalance), "0.00")
    
    '''        If Rec!tnySyncFlag Is Null Then
    '''            vsGrid.TextMatrix(mRowCnt, 3) = "NOT SYNC TO WEB"
    '''        ElseIf Rec!tnySyncFlag = 1 Then
    '''            vsGrid.TextMatrix(mRowCnt, 3) = "SYNC TO WEB"
    '''        End If
    
            Rec.MoveNext
            mRowCnt = mRowCnt + 1
        Wend
        Rec.Close
    End Sub

    Private Sub UpdateDailyIndex()
    
        '==================================================================================='
        ' CHECK DAILY INDEX WITH Transactions AND UPDATE DailyIndex FOR ANY MISSING DATES
        '==================================================================================='
        
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQl  As String
        Dim mRowCnt As Integer
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQl = " DECLARE @intID AS INTEGER" & vbCrLf
        mSQl = mSQl + " SELECT @intID= isNull(MAX (intID),0) FROM faDailyIndex" & vbCrLf
        mSQl = mSQl + " INSERT INTO faDailyIndex" & vbCrLf
        mSQl = mSQl + " (intID, intFinYearID, dtDate, tnyExtractFlag, tnySyncFlag, tnyVerificationFlag, tnyResetRequestFlag)" & vbCrLf
        mSQl = mSQl + " SELECT  -1,intFinancialYearID,dtTransactionDate,1,NULL,NULL,NULL" & vbCrLf
        mSQl = mSQl + "     From faDailyIndex" & vbCrLf
        mSQl = mSQl + "     RIGHT  JOIN faTransactions ON faDailyIndex.dtDate =  faTransactions.dtTransactionDate" & vbCrLf
        mSQl = mSQl + "     Where IsNull(faTransactions.tnyStatus, 0) <> 4" & vbCrLf
        mSQl = mSQl + " GROUP BY dtTransactionDate,intFinancialYearID" & vbCrLf
        mSQl = mSQl + " ORDER BY dtTransactionDate" & vbCrLf
        mSQl = mSQl + " UPDATE faDailyIndex SET @intID = intID = @intID + 1 WHERE ISNULL(intID,0)=-1" & vbCrLf
        
        objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
        
        'Call FillGrid
        
        Call CheckCashBalnceBalances
        
        
        mCnn.Close
    
    End Sub

    Private Sub CheckCashBalnceBalances()
    
        '==================================================================================='
        ' COMPARE BANK AND CASH BALANCES IN DAILYEXTRACTS WITH TRANSACTIONCHILD
        '==================================================================================='
        
        Dim mCnn                   As New ADODB.Connection
        Dim objDB                  As New clsDB
        Dim RecDailyExtracts       As New ADODB.Recordset
        Dim RecTransactionChild    As New ADODB.Recordset
        Dim mSQl                   As String
        Dim mSQLDailyExtracts      As String
        Dim mSQLTransactionChild   As String
        Dim mTotRowDailyExtracts   As Integer
        Dim mTotRowTranChild       As Integer
        Dim mReExtract             As Boolean
        
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQLDailyExtracts = " SELECT A.dtDate,SUM(CASHBALANCE) CASHBALANCE,SUM(BANKBALANCE) BANKBALANCE FROM"
        mSQLDailyExtracts = mSQLDailyExtracts + "     ("
        mSQLDailyExtracts = mSQLDailyExtracts + "     SELECT  faDailyExtracts.dtDate dtDate,"
        mSQLDailyExtracts = mSQLDailyExtracts + "             CASE WHEN faAccountHeads.intGroupID=1 Then"
        mSQLDailyExtracts = mSQLDailyExtracts + "             SUM(fltAmount) ELSE 0 END CASHBALANCE,"
        mSQLDailyExtracts = mSQLDailyExtracts + "             CASE WHEN faAccountHeads.intGroupID=2 Then"
        mSQLDailyExtracts = mSQLDailyExtracts + "             SUM(fltAmount) ELSE 0 END BANKBALANCE"
        mSQLDailyExtracts = mSQLDailyExtracts + "     From faDailyExtracts"
        mSQLDailyExtracts = mSQLDailyExtracts + "     INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faDailyExtracts.intAccountHeadID"
        mSQLDailyExtracts = mSQLDailyExtracts + "     WHERE faDailyExtracts.intFinancialYearID= " & mYearID & ""
        mSQLDailyExtracts = mSQLDailyExtracts + "           AND intMonthID= " & mMonthID & ""
        mSQLDailyExtracts = mSQLDailyExtracts + "           AND faAccountHeads.intGroupID in (1,2)"
        mSQLDailyExtracts = mSQLDailyExtracts + "     GROUP BY faDailyExtracts.dtDate,faAccountHeads.intGroupID"
        mSQLDailyExtracts = mSQLDailyExtracts + "    )A"
        mSQLDailyExtracts = mSQLDailyExtracts + "    GROUP BY A.dtDate"
        mSQLDailyExtracts = mSQLDailyExtracts + "    ORDER BY A.dtDate"
        
        RecDailyExtracts.CursorLocation = adUseClient
        RecDailyExtracts.Open mSQLDailyExtracts, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        If Not (RecDailyExtracts.EOF And RecDailyExtracts.BOF) Then
            mTotRowDailyExtracts = RecDailyExtracts.RecordCount
        End If
        
        mSQLTransactionChild = "SELECT A.dtDate,SUM(CASHBALANCE) CASHBALANCE,SUM(BANKBALANCE) BANKBALANCE FROM"
        mSQLTransactionChild = mSQLTransactionChild + "        ("
        mSQLTransactionChild = mSQLTransactionChild + "     SELECT  dtTransactionDate dtDate,"
        mSQLTransactionChild = mSQLTransactionChild + "             CASE WHEN faAccountHeads.intGroupID=1 Then"
        mSQLTransactionChild = mSQLTransactionChild + "             SUM((tinDebitOrCreditFlag * -2 + 1) * -1 * faTransactionChild.fltAmount) ELSE 0 END CASHBALANCE,"
        mSQLTransactionChild = mSQLTransactionChild + "             CASE WHEN faAccountHeads.intGroupID=2 Then"
        mSQLTransactionChild = mSQLTransactionChild + "             SUM((tinDebitOrCreditFlag * -2 + 1) * -1 * faTransactionChild.fltAmount) ELSE 0 END BANKBALANCE"
        mSQLTransactionChild = mSQLTransactionChild + "     From faTransactionChild"
        mSQLTransactionChild = mSQLTransactionChild + "     INNER JOIN faTransactions ON faTransactionChild.intTransactionID=faTransactions.intTransactionID"
        mSQLTransactionChild = mSQLTransactionChild + "     INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faTransactionChild.intAccountHeadID"
        mSQLTransactionChild = mSQLTransactionChild + "     WHERE faTransactions.intFinancialYearID= " & mYearID & ""
        mSQLTransactionChild = mSQLTransactionChild + "           AND MONTH(dtTransactionDate)= " & mMonthID & ""
        mSQLTransactionChild = mSQLTransactionChild + "           AND IsNull(faTransactions.tnyStatus, 0) <> 4"
        mSQLTransactionChild = mSQLTransactionChild + "           AND faAccountHeads.intGroupID in (1,2)"
        mSQLTransactionChild = mSQLTransactionChild + "     GROUP BY faTransactions.dtTransactionDate,faAccountHeads.intGroupID"
        mSQLTransactionChild = mSQLTransactionChild + "     )A"
        mSQLTransactionChild = mSQLTransactionChild + "     GROUP BY A.dtDate"
        mSQLTransactionChild = mSQLTransactionChild + "     ORDER BY A.dtDate"
        
        RecTransactionChild.CursorLocation = adUseClient
        RecTransactionChild.Open mSQLTransactionChild, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        If Not (RecTransactionChild.EOF And RecTransactionChild.BOF) Then
            mTotRowTranChild = RecTransactionChild.RecordCount
        End If
        
            If mTotRowDailyExtracts <> mTotRowTranChild Then
                'Re-Extracts
                
                mSQl = "EXEC spDailyExtractByYear " & mYearID & ""
                objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
                
                mSQl = ""
                mSQl = "UPDATE faDailyIndex SET tnyExtractFlag=NULL, tnySyncFlag=NULL WHERE intFinYearID=" & mYearID & " "
                objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
                Exit Sub
            Else
                While Not (RecTransactionChild.EOF Or RecTransactionChild.BOF)
                  If RecDailyExtracts!dtDate = RecTransactionChild!dtDate Then
                     If RecDailyExtracts!CASHBALANCE <> RecTransactionChild!CASHBALANCE And RecDailyExtracts!BankBalance <> RecTransactionChild!BankBalance Then
                        mReExtract = True
                     End If
                  End If
                  RecDailyExtracts.MoveNext
                  RecTransactionChild.MoveNext
                Wend
            End If
       
        
        If mReExtract = True Then
            mSQl = "EXEC spDailyExtractByYear " & mYearID & ""
            objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
            
            mSQl = ""
            mSQl = "UPDATE faDailyIndex SET tnyExtractFlag=NULL, tnySyncFlag=NULL WHERE intFinYearID=" & mYearID & " "
            objDB.ExecuteSP mSQl, , , , mCnn, adCmdText
            
        End If
        mCnn.Close
        
    End Sub
  
