VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSourceFundSplitUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Source"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmSourceFundSplitUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAllotmentReceived 
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtOthers 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtTreasuryBalance 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   1500
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3990
      Width           =   1335
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6135
      _cx             =   10821
      _cy             =   4260
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
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSourceFundSplitUp.frx":1CCA
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Other Funds"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Treasury Balance"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   4
      Top             =   4080
      Width           =   375
   End
End
Attribute VB_Name = "frmSourceFundSplitUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mTreasuryID As Integer
    Dim mLoASource As Integer
    Dim mLoACategory As Integer
    Dim mSaveMode As Integer
   
    Private Sub cmdSave_Click()
        If val(txtTotal.Text) = val(txtTreasuryBalance.Text) Then
            Call SaveAllotmentLetters(mTreasuryID)
        Else
            MsgBox " Amount Miss match in Treasury With Source", vbInformation
            Exit Sub
        End If
    End Sub
    Private Sub SaveAllotmentLetters(mTreasuryID As Integer)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mArrIN As Variant
        Dim mArrOut As Variant
        Dim objDB As New clsDB
        Dim mCnt As Integer
        Dim objAcc As New clsAccounts
        Dim mCheckSave As Boolean
        Dim mAllotmentSourceID As Integer
        Dim mAllotmentDevSourceID As Integer
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
    For mCnt = 1 To vsGrid.Rows - 1
        If val(vsGrid.TextMatrix(mCnt, 2)) <> 0 Then
'             msql = "SELECT * FROM faAllotmentLetters WHERE  ISNULL(tnyGroupID,0)=30 AND intSourceOfFundID= " & val(vsGrid.TextMatrix(mCnt, 3)) & " AND intCategoryID=" & val(vsGrid.TextMatrix(mCnt, 4)) & ""
'             Rec.Open msql, mCnn
'             If Not (Rec.EOF And Rec.BOF) Then
'                 mCheckSave = False
'                 'Exit Sub
'             Else
'                 mCheckSave = True
'             End If
'             Rec.Close
'
'             If mCheckSave = True Then
                     mArrIN = Array(-1, _
                                     Null, _
                                     gbTransactionDate, _
                                     Null, _
                                     val(vsGrid.TextMatrix(mCnt, 3)), val(vsGrid.TextMatrix(mCnt, 4)), _
                                     Null, Null, _
                                     Null, _
                                     Null, _
                                     Null, _
                                     Null, _
                                     Null, _
                                     Null, _
                                     -1 * val(vsGrid.TextMatrix(mCnt, 2)), _
                                     Null, _
                                     Null, _
                                     Null, _
                                     Null, _
                                     gbUserID, _
                                     gbTransactionDate, _
                                     Null, _
                                     gbLocalBodyID, _
                                     gbFinancialYearID, _
                                     1, _
                                     0, Null, Null, 30, IIf(gbLBPanchayat = 1, 4010, 4006), 0, Null _
                                 )
                     objDB.ExecuteSP "spSaveAllotmentLetter", mArrIN, mArrOut, , mCnn, adCmdStoredProc
                     mAllotmentSourceID = mArrOut(0, 0)
                End If
'             End If
             Next
'             If mCheckSave = True Then
                If val(txtOthers.Text) <> 0 Then 'ANY OTHER FUNDS
                    mArrIN = Array(-1, _
                                        Null, _
                                        gbTransactionDate, _
                                        Null, _
                                        Null, Null, _
                                        Null, Null, _
                                        Null, _
                                        Null, _
                                        Null, _
                                        Null, _
                                        Null, _
                                        Null, _
                                        val(txtOthers.Text), _
                                        Null, _
                                        Null, _
                                        Null, _
                                        Null, _
                                        gbUserID, _
                                        gbTransactionDate, _
                                        Null, _
                                        gbLocalBodyID, _
                                        gbFinancialYearID, _
                                        1, _
                                        0, Null, Null, 40, IIf(gbLBPanchayat = 1, 4010, 4006), 0, Null _
                                    )
                        objDB.ExecuteSP "spSaveAllotmentLetter", mArrIN, , , mCnn, adCmdStoredProc
                End If
                    'DEVELOPMENT FUND GENERAL
                    mArrIN = Array(-1, _
                                    Null, _
                                    gbTransactionDate, _
                                    Null, _
                                    1, 1, _
                                    Null, Null, _
                                    val(txtTreasuryBalance.Tag), _
                                    Null, _
                                    Null, _
                                    Null, _
                                    Null, _
                                    Null, _
                                    val(txtTotal.Text), _
                                    Null, _
                                    Null, _
                                    Null, _
                                    Null, _
                                    gbUserID, _
                                    gbTransactionDate, _
                                    Null, _
                                    gbLocalBodyID, _
                                    gbFinancialYearID, _
                                    1, _
                                    0, Null, Null, 30, IIf(gbLBPanchayat = 1, 4010, 4006), 0, Null _
                                )
                    objDB.ExecuteSP "spSaveAllotmentLetter", mArrIN, mArrOut, , mCnn, adCmdStoredProc
                    mAllotmentDevSourceID = mArrOut(0, 0)
'            End If
        'End If
    
    
    If gbLBPanchayat = 1 Then
        frmContraEntry.vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount2

        Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount2)
    Else
        frmContraEntry.vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount2
        Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount6)
    End If
    frmContraEntry.vsGrid.TextMatrix(1, 2) = objAcc.AccountHead
    frmContraEntry.vsGrid.TextMatrix(1, 4) = val(txtTotal.Text)
    frmContraEntry.cmbSource.Tag = mAllotmentSourceID
    frmContraEntry.cmbCategory.Tag = mAllotmentDevSourceID
    
    mCnn.Close
    Unload Me
    End Sub
    Public Property Let TreasuryID(mData As Long)
        mTreasuryID = mData
    End Property
    Public Property Get TreasuryID() As Long
        TreasuryID = mTreasuryID
    End Property
    Private Sub Form_Load()
         mSaveMode = 0
        Call mtreasuryBalance(mTreasuryID)
        Call Fillgrid(mTreasuryID)
    End Sub
    Private Sub mtreasuryBalance(mTreasuryID As Integer)
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim objDB As New clsDB
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT faTransactionChild.intAccountHeadID,vchAccountHeadCode,vchBankName,isnull(sum(fltAmount*((faTransactionChild.tinDebitOrCreditFlag*2)-1)),0) fltAmount"
        mSql = mSql + " From faTransactionChild"
        mSql = mSql + " INNER  JOIN faTransactions ON faTransactions.intTransactionID=faTransactionChild.intTransactionID"
        mSql = mSql + " INNER  JOIN  faAccountHeads ON faTransactionChild.intAccountheadID=faAccountHeads.intAccountheadID"
        mSql = mSql + " INNER JOIN faBAnks ON faBanks.intAccountHeadID=faAccountHeads.intAccountheadID"
        mSql = mSql + " Where  faBanks.intAccountHeadID = " & mTreasuryID
        mSql = mSql + " AND faAccountHeads.intGroupID=2 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL)"
        mSql = mSql + " Group by vchBankName,vchAccountHeadCode,faTransactionChild.intAccountHeadID"
        
        'dtTransactionDate <= '31/Mar/2015' AND
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            txtTreasuryBalance.Text = Rec!fltAmount
            txtTreasuryBalance.Tag = mTreasuryID
            txtTreasuryBalance.Enabled = False
        End If
        txtTreasuryBalance.Enabled = False
        Rec.Close
        mCnn.Close
    End Sub
    Private Sub Fillgrid(mTreasuryID As Integer)
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        Dim mArrIN As Variant
        Dim objDB As New clsDB
        Dim mSource As String
        Dim mCategory As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If gbLBPanchayat = 1 Then
            If mTreasuryID = 1418 Then
                mSource = " (4,31,32,33,34,35)"
            ElseIf mTreasuryID = 1490 Then 'DF(General)
                mSource = "   (1,2,10,11,12,13,14,27,28,21)"
                mCategory = " AND intFundCategoryID=1 "
            ElseIf mTreasuryID = 1494 Then 'SCP
                mSource = "    (29,10,11,12,13,14) "
                mCategory = " AND intFundCategoryID=2 "
            ElseIf mTreasuryID = 1495 Then 'TSP
                mSource = "     (30,10,11,12,13,14) "
                mCategory = " AND intFundCategoryID=3 "
            ElseIf mTreasuryID = 1491 Then 'Maintainance
                mSource = "     (16,17)"
            ElseIf mTreasuryID = 1492 Then 'CFC-Award Grant
                mSource = "     (25)"
            ElseIf mTreasuryID = 1493 Then 'KLGSDP Grant
                mSource = "  (26)"
            Else
                Exit Sub
            End If
        Else
             If mTreasuryID = 1512 Then
                mSource = " (4,31,32,33,34,35)"
            ElseIf mTreasuryID = 1535 Then  'DF(General)
                mSource = " (1,2,10,11,12,13,14,27,28,21) "
                mCategory = " AND intFundCategoryID=1 "
            ElseIf mTreasuryID = 1816 Then  'SCP
                mSource = "  (29,10,11,12,13,14) "
                mCategory = " AND intFundCategoryID=2 "
            ElseIf mTreasuryID = 1817 Then  'TSP
                mSource = "  (30,10,11,12,13,14) "
                mCategory = " AND intFundCategoryID=3 "
            ElseIf mTreasuryID = 1539 Then  'Maintainance
                mSource = " (16,17)"
            ElseIf mTreasuryID = 1755 Then  'CFC-Award Grant
                mSource = " (25)"
            ElseIf mTreasuryID = 1756 Then  'KLGSDP Grant
                mSource = " (26)"
            Else
                Exit Sub
            End If
        End If
        
        
        mSql = " select ISNULL(A.fltRequestedAmt,0) fltRequestedAmt,suSourceOfFund.intSourceFundID,suSourceOfFund.vchSourceFundName FROM"
        mSql = mSql + " ("
        mSql = mSql + " SELECT  SUM(fltRequestedAmt) fltRequestedAmt ,intSourceID "
        mSql = mSql + " From faAllotments"
        mSql = mSql + " INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotments.intSourceID"
        mSql = mSql + " WHERE faAllotments.intFinancialYearID = 2015 AND ISNULL(faAllotments.intTreasuryID,0)=0 AND ISNULL(tnyStatus,0)<>2 AND suSourceOfFund.intSourceFundID IN " & mSource
        mSql = mSql + " " & mCategory & " "
        mSql = mSql + " GROUP BY intSourceID"
        mSql = mSql + " )A"
        mSql = mSql + " RIGHT OUTER JOIN   suSourceOfFund on suSourceOfFund.intSourceFundID=A.intSourceID"
        mSql = mSql + " WHERE  suSourceOfFund.intSourceFundID IN  " & mSource
                
        Rec.Open mSql, mCnn
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRow = 1
        If Not (Rec.BOF And Rec.EOF) Then
             While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                If gbLBPanchayat = 1 Then
                    If mTreasuryID = 1494 Then
                        vsGrid.TextMatrix(mRow, 1) = "SCP"
                        vsGrid.TextMatrix(mRow, 4) = 2
                    ElseIf mTreasuryID = 1495 Then
                        vsGrid.TextMatrix(mRow, 1) = "TSP"
                        vsGrid.TextMatrix(mRow, 4) = 3
                    
                    Else
                        vsGrid.TextMatrix(mRow, 1) = "GENERAL"
                        vsGrid.TextMatrix(mRow, 4) = 1
                    End If
                Else
                    If mTreasuryID = 1816 Then
                        vsGrid.TextMatrix(mRow, 1) = "SCP"
                        vsGrid.TextMatrix(mRow, 4) = 2
                    ElseIf mTreasuryID = 1817 Then
                        vsGrid.TextMatrix(mRow, 1) = "TSP"
                        vsGrid.TextMatrix(mRow, 4) = 3
                    
                    Else
                        vsGrid.TextMatrix(mRow, 1) = "GENERAL"
                        vsGrid.TextMatrix(mRow, 4) = 1
                    End If
                End If
                vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!fltRequestedAmt), 0, Rec!fltRequestedAmt)
                vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
                mRow = mRow + 1
                Rec.MoveNext
             Wend
        End If
                

        Call GetBalance
        Call Calculate
        Rec.Close
        mCnn.Close
    End Sub
    Private Sub GetBalance()
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        Dim mArrIN As Variant
        Dim objDB As New clsDB
        Dim mCount As Integer
        
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            For mCount = 1 To vsGrid.Rows - 1
                mArrIN = Array(val(vsGrid.TextMatrix(mCount, 3)), val(vsGrid.TextMatrix(mCount, 4)))
                Set Rec = objDB.ExecuteSP("spGetAllotmentReceived", mArrIN, , , mCnn, adCmdStoredProc)
                
                If Not (Rec.BOF And Rec.EOF) Then
                    'MsgBox Rec!intSourceOfFundID
                    If val(vsGrid.TextMatrix(mCount, 3)) = Rec!intSourceOfFundID Then 'val(vsGrid.TextMatrix(mCount, 5)) <> 0 Then
                        vsGrid.TextMatrix(mCount, 2) = val(Rec!AmountReceived) - val(vsGrid.TextMatrix(mCount, 5))
            
                    End If
                End If
            Next
            Rec.Close
        mCnn.Close
    End Sub
    Private Sub Calculate()
        Dim mTotal       As Double
        Dim mCount       As Integer
        
        txtTotal.Text = ""
        For mCount = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mCount, 2)) <> 0 Then
                mTotal = mTotal + Format(val(vsGrid.TextMatrix(mCount, 2)), "0.00")
                txtTotal.Text = Format(mTotal, "0.00")
            End If
        Next
        If val(txtOthers.Text) <> 0 Then
            txtTotal.Text = val(txtTotal.Text) + val(txtOthers.Text)
        End If
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
''''        Dim objAcc As New clsAccounts
''''        If mSaveMode = 1 Then
''''            If gbLBPanchayat = 1 Then
''''                frmContraEntry.vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount2
''''            Else
''''                frmContraEntry.vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount6
''''            End If
''''            If gbLBPanchayat = 1 Then
''''                Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount2)
''''            Else
''''                Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount6)
''''            End If
''''
''''
''''            frmContraEntry.vsGrid.TextMatrix(1, 2) = objAcc.AccountHead
''''            frmContraEntry.vsGrid.TextMatrix(1, 4) = val(txtTotal.Text)
''''        End If
    End Sub
    Private Sub txtOthers_LostFocus()
        Call Calculate
    End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Call Calculate
    End Sub

