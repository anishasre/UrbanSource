VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCBSourceofFundTreasury 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Closing Balance Of Treasury (Source Of Fund )"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCBSourceofFundTreasury.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   13680
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   11775
      Begin VB.CommandButton cmdApprove 
         Caption         =   "APPROVE CLOSING  BALANCE"
         Height          =   615
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdListVrs 
         Caption         =   "LIST OF RECEIPTS"
         Height          =   615
         Left            =   9000
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "VERIFY CLOSING  BALANCE"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   11880
      TabIndex        =   4
      Top             =   720
      Width           =   2535
      Begin VB.Label lbl 
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
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblmsg 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame frmNewProcess 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11775
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   9960
         TabIndex        =   10
         Top             =   4680
         Width           =   1575
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridSource 
         Height          =   4455
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   7858
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCBSourceofFundTreasury.frx":1CCA
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
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
         _cx             =   11033
         _cy             =   6800
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCBSourceofFundTreasury.frx":1D68
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
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14535
      TabIndex        =   0
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "frmCBSourceofFundTreasury"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApprove_Click()
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mRow As Integer
    Dim mSql As String
    Dim mArrIN As Variant

    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
     mSql = "Update faBAnkSource set tnyClosingFlag = 2 WHERE tnyClosingFlag <> 9  "
    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
    
    mSql = "Update faBankSource SET tnyClosingFlag = 9  WHERE fltClosingBalance = 0 "
    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
    
    cmdVerify.Enabled = False
    cmdApprove.Enabled = False
    mCnn.Close
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub FillGridSource(mAccountHeadID As Integer)
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCnt As Integer
    
    On Error GoTo err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = " SELECT intSourceFundID,vchSourceFundName,fltAmount"
        mSql = mSql + " From faExtractAllotments"
        mSql = mSql + " INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faExtractAllotments.intSourceOfFundID"
        mSql = mSql + " Where intFinancialYearID = 2015"

        If gbLBPanchayat = 1 Then
            If mAccountHeadID = 1418 Then
                mSql = mSql + " AND intSourceOfFundID IN (4,31,32,33,34,35)"
            ElseIf mAccountHeadID = 1490 Then 'DF(General)
                mSql = mSql + " AND intSourceOfFundID IN (1,2,10,11,12,13,14,27,28,21) AND intCategoryID=1"
            ElseIf mAccountHeadID = 1494 Then 'SCP
                mSql = mSql + " AND intSourceOfFundID IN (29,2,10,11,12,13,14) AND intCategoryID=2"
            ElseIf mAccountHeadID = 1495 Then 'TSP
                mSql = mSql + " AND intSourceOfFundID IN (30,2,10,11,12,13,14) AND intCategoryID=3"
            ElseIf mAccountHeadID = 1491 Then 'Maintainance
                mSql = mSql + " AND intSourceOfFundID IN (16,17)"
            ElseIf mAccountHeadID = 1492 Then 'CFC-Award Grant
                mSql = mSql + " AND intSourceOfFundID IN (25)"
            ElseIf mAccountHeadID = 1493 Then 'KLGSDP Grant
                mSql = mSql + " AND intSourceOfFundID IN (26)"
            Else
                Exit Sub
            End If
        Else
             If mAccountHeadID = 1512 Then
                mSql = mSql + " AND intSourceOfFundID IN (4,31,32,33,34,35)"
            ElseIf mAccountHeadID = 1535 Then  'DF(General)
                mSql = mSql + " AND intSourceOfFundID IN (1,2,10,11,12,13,14,27,28,21) AND intCategoryID=1"
            ElseIf mAccountHeadID = 1816 Then  'SCP
                mSql = mSql + " AND intSourceOfFundID IN (29,2,10,11,12,13,14) AND intCategoryID=2"
            ElseIf mAccountHeadID = 1817 Then  'TSP
                mSql = mSql + " AND intSourceOfFundID IN (30,2,10,11,12,13,14) AND intCategoryID=3"
            ElseIf mAccountHeadID = 1539 Then  'Maintainance
                mSql = mSql + " AND intSourceOfFundID IN (16,17)"
            ElseIf mAccountHeadID = 1755 Then  'CFC-Award Grant
                mSql = mSql + " AND intSourceOfFundID IN (25)"
            ElseIf mAccountHeadID = 1756 Then  'KLGSDP Grant
                mSql = mSql + " AND intSourceOfFundID IN (26)"
            Else
                Exit Sub
            End If
        End If
        Rec.CursorLocation = adUseClient
        
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        vsGridSource.Clear 1, 1
        vsGridSource.Rows = 1
        While Not (Rec.EOF Or Rec.BOF)
            vsGridSource.Rows = vsGridSource.Rows + 1
            vsGridSource.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            vsGridSource.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            vsGridSource.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
            vsGridSource.TextMatrix(mRowCnt, 3) = IIf(IsNull(mAccountHeadID), "", mAccountHeadID)
            Rec.MoveNext
            mRowCnt = mRowCnt + 1
        Wend
        
        Rec.Close
        mCnn.Close
        Call CalculateTotal
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillGridExtract()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCnt As Integer

    
    On Error GoTo err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSql = " SELECT faTransactionChild.intAccountHeadID,vchAccountHeadCode,vchBankName,isnull(sum(fltAmount*((faTransactionChild.tinDebitOrCreditFlag*2)-1)),0) fltAmount "
    mSql = mSql + " From faTransactionChild"
    mSql = mSql + " INNER  JOIN faTransactions ON faTransactions.intTransactionID=faTransactionChild.intTransactionID"
    mSql = mSql + " INNER  JOIN  faAccountHeads ON faTransactionChild.intAccountheadID=faAccountHeads.intAccountheadID"
    mSql = mSql + " INNER JOIN faBAnks ON faBanks.intAccountHeadID=faAccountHeads.intAccountheadID"
    mSql = mSql + " WHERE   dtTransactionDate < = '31-Mar-2015' AND ( vchAccountHeadCode LIKE '450250%' OR vchAccountHeadCode LIKE '450650%') "
    'mSql = mSql + " WHERE   dtTransactionDate < =' 31 / Mar / 2015 ' AND vchAccountHeadCode LIKE '450250%' OR vchAccountHeadCode LIKE '450650%'"    '''' ::::Changed By Aiby 08-Jul-2015
    mSql = mSql + " AND faAccountHeads.intGroupID=2 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL)"
    mSql = mSql + " Group by vchBankName,vchAccountHeadCode,faTransactionChild.intAccountHeadID"
    
    Rec.CursorLocation = adUseClient
    
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    
    Rec.Close
    objDB.ExecuteSP "spExtractBankSource", , , , mCnn 'EXTRACTED
    mCnn.Close
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCnt As Integer

    
    On Error GoTo err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSql = " SELECT * FROM faBankSource"
    mSql = mSql + " INNER JOIN faAccountHeads ON faBankSource.intBankID=faAccountHeads.intAccountHeadID"
    
    Rec.CursorLocation = adUseClient
    
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltClosingBalance), "", Rec!fltClosingBalance)
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        If Rec!tnyClosingFlag = 2 Or Rec!tnyClosingFlag = 9 Then
            vsGrid.TextMatrix(mRowCnt, 3) = vbChecked
        Else
             vsGrid.TextMatrix(mRowCnt, 3) = vbUnchecked
        End If
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close
    mCnn.Close
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmdVerify_Click()
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mRow As Integer
    Dim mSql As String
    Dim mArrIN As Variant
    Dim mLoop As Integer
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
     mSql = "Update faBAnkSource set tnyClosingFlag=1 WHERE tnyClosingFlag<>9 "
    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
    cmdVerify.Enabled = False
    
    For mLoop = 1 To vsGrid.Rows - 1
        mArrIN = Array(val(vsGrid.TextMatrix(mLoop, 4)))
        objDB.ExecuteSP "spSaveBankSourceChild", mArrIN, , , mCnn
    Next
    
    mCnn.Close
End Sub

Private Sub Form_Load()
    Dim mExtractedStatus As Integer
    Dim mMsg As String
    Dim mLoop As Integer
    Dim mVerifyStatus As Integer
    
    mExtractedStatus = GetStatusFlag()
    If mExtractedStatus <> 2 Then
        cmdVerify.Enabled = False
        cmdApprove.Enabled = False
        mMsg = ""
        mMsg = mMsg + " Closing Balance Of Source Of Fund is not Approved" + vbCrLf
        mMsg = mMsg + " (Utility>>Annual Financial Statements-Finalization>>)"
        MsgBox mMsg, vbInformation
        Exit Sub
    End If
    mVerifyStatus = checkVerify()
    If mVerifyStatus < 0 Then
        Call FillGridExtract
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdVerify.Enabled = True
            cmdApprove.Enabled = False
        Else
            cmdVerify.Enabled = False
            cmdApprove.Enabled = False
        End If
    ElseIf mVerifyStatus = 0 Then
        Call FillGrid
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdVerify.Enabled = True
            cmdApprove.Enabled = False
        Else
            cmdVerify.Enabled = False
            cmdApprove.Enabled = False
        End If
    ElseIf mVerifyStatus = 1 Then
         Call FillGrid
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdVerify.Enabled = False
            cmdApprove.Enabled = False
        Else
            cmdVerify.Enabled = False
            cmdApprove.Enabled = True
        End If
    Else
        Call FillGrid
        cmdVerify.Enabled = False
        cmdApprove.Enabled = False
    End If
    lbl.Visible = True
   lblmsg.Caption = "Verify Balances Of Treasury with Source Of Fund As On 31 March 2015"
End Sub
Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub vsGrid_Click()
    If vsGrid.Row > 0 Then
        Call FillGridSource(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
    End If
End Sub
Private Sub CalculateTotal()
    Dim mTotal As Double
    Dim mLoop As Integer

    mTotal = 0
    For mLoop = 1 To vsGridSource.Rows - 1
        mTotal = mTotal + val(vsGridSource.TextMatrix(mLoop, 1))
    Next
    txtTotal.Text = mTotal
End Sub

Private Function GetStatusFlag() As Integer
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mTrAccHeadId As Integer
    
    If objDB.SetConnection(mCnn) Then
        mSql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            GetStatusFlag = Rec!tnyStatus
        Else
            GetStatusFlag = -1
        End If
        Rec.Close
    End If
End Function
Private Function checkVerify() As Integer
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String

    
    If objDB.SetConnection(mCnn) Then
        mSql = "SELECT tnyClosingFlag FROM faBankSource "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            checkVerify = Rec!tnyClosingFlag
        Else
            checkVerify = -1
        End If
        Rec.Close
    End If
End Function
